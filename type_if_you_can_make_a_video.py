#!/usr/bin/env python3
"""
Type If You Can, Make A Video

Simple local-first video builder.

Workflow:
1. Type or paste plain text.
2. Click Make Slides.
3. Optionally choose a voice sample and background folder.
4. Export MP4 with Chatterbox-Turbo narration.

Optional features:
- faster-whisper transcription from imported or recorded audio
- Ollama rewrite helper
- one-click pip setup for common missing packages
- Chatterbox-Turbo warmup to download model weights
"""

from __future__ import annotations

import json
import math
import os
import platform
import re
import shutil
import subprocess
import sys
import textwrap
import threading
import time
import urllib.error
import urllib.request
import wave
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

try:
    from PIL import Image, ImageColor, ImageDraw, ImageFont, ImageOps, ImageTk
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False
    Image = ImageColor = ImageDraw = ImageFont = ImageOps = ImageTk = None

try:
    import sounddevice as sd
    SOUNDDEVICE_AVAILABLE = True
except Exception:
    SOUNDDEVICE_AVAILABLE = False
    sd = None

try:
    import soundfile as sf
    SOUNDFILE_AVAILABLE = True
except Exception:
    SOUNDFILE_AVAILABLE = False
    sf = None

try:
    from faster_whisper import WhisperModel
    FASTER_WHISPER_AVAILABLE = True
except Exception:
    FASTER_WHISPER_AVAILABLE = False
    WhisperModel = None

APP_TITLE = "Type If You Can, Make A Video"
PROJECT_EXTENSION = ".type2video.json"
PREVIEW_BG = "#111827"
TIMELINE_BG = "#0F172A"
TIMELINE_GRID = "#1E293B"
TIMELINE_TEXT = "#E5E7EB"
CLIP_FILL = "#2563EB"
CLIP_SELECTED = "#F59E0B"
CLIP_TEXT = "#F8FAFC"
TIMELINE_HEADER_WIDTH = 72
TIMELINE_LAYER_HEIGHT = 58
TIMELINE_MIN_DURATION = 0.5
PIXELS_PER_SECOND = 75
DEFAULT_WIDTH = 1280
DEFAULT_HEIGHT = 720
TITLE_FONT = ("Arial", 20, "bold")
BODY_FONT = ("Arial", 12)
RECOMMENDED_PYTHON = "3.11"


@dataclass
class TextStyle:
    title_font: str = "Arial"
    body_font: str = "Arial"
    title_size: int = 34
    body_size: int = 24
    title_align: str = "left"
    body_align: str = "left"
    fg_color: str = "#F8FAFC"
    overlay_color: str = "#0F172A"
    overlay_alpha: int = 185


@dataclass
class Slide:
    slide_id: str
    title: str = "Untitled Slide"
    start: float = 0.0
    duration: float = 5.0
    layer: int = 0
    body: str = ""
    narration: str = ""
    background: str = ""
    crop_x: float = 0.5
    crop_y: float = 0.5
    style: TextStyle = field(default_factory=TextStyle)

    @property
    def end(self) -> float:
        return self.start + self.duration

    def clone(self) -> "Slide":
        return Slide(
            slide_id=self.slide_id,
            title=self.title,
            start=self.start,
            duration=self.duration,
            layer=self.layer,
            body=self.body,
            narration=self.narration,
            background=self.background,
            crop_x=self.crop_x,
            crop_y=self.crop_y,
            style=TextStyle(**asdict(self.style)),
        )


@dataclass
class AppSettings:
    canvas_width: int = DEFAULT_WIDTH
    canvas_height: int = DEFAULT_HEIGHT
    max_chars_per_slide: int = 260
    min_chars_per_slide: int = 70
    voice_sample_wav: str = ""
    background_folder: str = ""
    ffmpeg_path: str = ""
    fps: int = 30
    gap_after_slide_sec: float = 0.25
    use_cuda: bool = True
    whisper_model: str = "small"
    whisper_device: str = "auto"
    ollama_model: str = "qwen3:8b"
    ollama_url: str = "http://localhost:11434/api/generate"
    chatterbox_exaggeration: float = 0.50
    chatterbox_cfg_weight: float = 0.50
    chatterbox_temperature: float = 0.80


@dataclass
class Project:
    name: str = "Untitled Project"
    raw_script: str = ""
    settings: AppSettings = field(default_factory=AppSettings)
    slides: List[Slide] = field(default_factory=list)

    def normalize_ids(self) -> None:
        seen = set()
        for i, slide in enumerate(self.slides, 1):
            if not slide.slide_id or slide.slide_id in seen:
                slide.slide_id = f"slide_{i:03d}"
            seen.add(slide.slide_id)

    def sort_slides(self) -> None:
        self.slides.sort(key=lambda s: (s.start, s.layer, s.slide_id))

    def total_duration(self) -> float:
        return max((s.end for s in self.slides), default=0.0)

    def max_layer(self) -> int:
        return max((s.layer for s in self.slides), default=0)


def project_to_dict(project: Project) -> dict:
    return {
        "name": project.name,
        "raw_script": project.raw_script,
        "settings": asdict(project.settings),
        "slides": [
            {
                "slide_id": s.slide_id,
                "title": s.title,
                "start": s.start,
                "duration": s.duration,
                "layer": s.layer,
                "body": s.body,
                "narration": s.narration,
                "background": s.background,
                "crop_x": s.crop_x,
                "crop_y": s.crop_y,
                "style": asdict(s.style),
            }
            for s in project.slides
        ],
    }


def project_from_dict(data: dict) -> Project:
    settings = AppSettings(**data.get("settings", {}))
    slides = []
    for item in data.get("slides", []):
        slides.append(
            Slide(
                slide_id=item.get("slide_id", ""),
                title=item.get("title", "Untitled Slide"),
                start=float(item.get("start", 0.0)),
                duration=float(item.get("duration", 5.0)),
                layer=int(item.get("layer", 0)),
                body=item.get("body", ""),
                narration=item.get("narration", ""),
                background=item.get("background", ""),
                crop_x=float(item.get("crop_x", 0.5)),
                crop_y=float(item.get("crop_y", 0.5)),
                style=TextStyle(**item.get("style", {})),
            )
        )
    project = Project(
        name=data.get("name", "Untitled Project"),
        raw_script=data.get("raw_script", ""),
        settings=settings,
        slides=slides,
    )
    project.normalize_ids()
    project.sort_slides()
    return project


def plain_project_source(project: Project) -> str:
    lines = [f"#project {project.name}", ""]
    for slide in sorted(project.slides, key=lambda s: (s.start, s.layer, s.slide_id)):
        attrs = (
            f'id="{slide.slide_id}" '
            f'title="{slide.title.replace(chr(34), chr(39))}" '
            f'start="{slide.start:.2f}" '
            f'duration="{slide.duration:.2f}" '
            f'layer="{slide.layer}" '
            f'background="{slide.background.replace(chr(34), chr(39))}" '
            f'crop_x="{slide.crop_x:.3f}" '
            f'crop_y="{slide.crop_y:.3f}"'
        )
        lines.append(f"[[slide {attrs}]]")
        lines.append("body:")
        lines.append(slide.body.rstrip())
        lines.append("")
        lines.append("narration:")
        lines.append(slide.narration.rstrip())
        lines.append("[[/slide]]")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


SECTION_RE = re.compile(r"^\s*(body|narration)\s*:\s*$", re.I)
SLIDE_OPEN_RE = re.compile(r"^\s*\[\[slide\s*(.*?)\]\]\s*$", re.I)
SLIDE_CLOSE_RE = re.compile(r"^\s*\[\[/slide\]\]\s*$", re.I)
ATTR_RE = re.compile(r'(\w+)\s*=\s*"(.*?)"')


def parse_attrs(attr_text: str) -> dict:
    out = {}
    for key, value in ATTR_RE.findall(attr_text):
        out[key.lower()] = value
    return out


def parse_plain_project_source(text: str, existing_settings: AppSettings, existing_raw: str) -> Project:
    project = Project(name="Untitled Project", raw_script=existing_raw, settings=existing_settings, slides=[])
    current = None
    section = None
    body_lines: List[str] = []
    narration_lines: List[str] = []

    def flush():
        nonlocal current, section, body_lines, narration_lines
        if current is None:
            return
        current.body = "\n".join(body_lines).strip()
        current.narration = "\n".join(narration_lines).strip()
        project.slides.append(current)
        current = None
        section = None
        body_lines = []
        narration_lines = []

    for line in text.splitlines():
        stripped = line.strip()
        if current is None:
            if stripped.startswith("#project"):
                parts = stripped.split(maxsplit=1)
                if len(parts) == 2:
                    project.name = parts[1].strip()
            else:
                m = SLIDE_OPEN_RE.match(line)
                if m:
                    attrs = parse_attrs(m.group(1))
                    current = Slide(
                        slide_id=attrs.get("id", ""),
                        title=attrs.get("title", "Untitled Slide"),
                        start=float(attrs.get("start", "0") or 0),
                        duration=max(TIMELINE_MIN_DURATION, float(attrs.get("duration", "5") or 5)),
                        layer=max(0, int(attrs.get("layer", "0") or 0)),
                        background=attrs.get("background", ""),
                        crop_x=min(1.0, max(0.0, float(attrs.get("crop_x", "0.5") or 0.5))),
                        crop_y=min(1.0, max(0.0, float(attrs.get("crop_y", "0.5") or 0.5))),
                    )
        else:
            if SLIDE_CLOSE_RE.match(line):
                flush()
            else:
                sec = SECTION_RE.match(line)
                if sec:
                    section = sec.group(1).lower()
                else:
                    if section == "body":
                        body_lines.append(line)
                    elif section == "narration":
                        narration_lines.append(line)
    flush()
    project.normalize_ids()
    project.sort_slides()
    return project


def normalize_newlines(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n")


def split_sentences_preserving_words(text: str) -> List[str]:
    text = normalize_newlines(text).strip()
    if not text:
        return []
    text = re.sub(r"\n{2,}", "\n\n", text)
    blocks = []
    for para in text.split("\n\n"):
        para = re.sub(r"\s+", " ", para.strip())
        if not para:
            continue
        pieces = re.split(r"(?<=[\.\!\?;:])\s+", para)
        for piece in pieces:
            piece = piece.strip()
            if piece:
                blocks.append(piece)
    return blocks


def chunk_long_text_by_words(text: str, max_chars: int) -> List[str]:
    words = text.split()
    chunks: List[str] = []
    current = []
    current_len = 0
    for word in words:
        needed = len(word) + (1 if current else 0)
        if current and current_len + needed > max_chars:
            chunks.append(" ".join(current).strip())
            current = [word]
            current_len = len(word)
        else:
            current.append(word)
            current_len += needed
    if current:
        chunks.append(" ".join(current).strip())
    return chunks


def title_from_chunk(text: str, index: int) -> str:
    words = re.findall(r"\w+(?:['-]\w+)?", text)
    if not words:
        return f"Slide {index}"
    title = " ".join(words[:7]).strip()
    if len(title) > 55:
        title = title[:52].rstrip() + "..."
    return title.title()


def estimate_duration_from_text(text: str) -> float:
    words = max(1, len(re.findall(r"\w+", text or "")))
    return round(max(2.3, words / 2.45), 2)


def slides_from_raw_text(raw_text: str, settings: AppSettings, template_style: Optional[TextStyle] = None) -> List[Slide]:
    max_chars = max(80, int(settings.max_chars_per_slide or 260))
    min_chars = max(30, int(settings.min_chars_per_slide or 70))
    segments = split_sentences_preserving_words(raw_text)
    if not segments:
        return []

    prepared: List[str] = []
    for seg in segments:
        if len(seg) > max_chars:
            prepared.extend(chunk_long_text_by_words(seg, max_chars))
        else:
            prepared.append(seg)

    chunks: List[str] = []
    current = ""
    for seg in prepared:
        proposal = (current + " " + seg).strip() if current else seg
        if current and len(proposal) > max_chars:
            chunks.append(current)
            current = seg
        else:
            current = proposal
            if len(current) >= min_chars and current.endswith((".", "?", "!", ";", ":")):
                chunks.append(current)
                current = ""
    if current:
        if chunks and len(current) < min_chars:
            chunks[-1] = (chunks[-1] + " " + current).strip()
        else:
            chunks.append(current)

    slides: List[Slide] = []
    start = 0.0
    style = template_style or TextStyle()
    for idx, chunk in enumerate(chunks, 1):
        slide = Slide(
            slide_id=f"slide_{idx:03d}",
            title=title_from_chunk(chunk, idx),
            start=round(start, 2),
            duration=estimate_duration_from_text(chunk),
            layer=0,
            body=chunk.strip(),
            narration=chunk.strip(),
            style=TextStyle(**asdict(style)),
        )
        slides.append(slide)
        start += slide.duration
    return slides


def find_font_file(font_name: str) -> Optional[str]:
    if not font_name:
        return None
    if os.path.isfile(font_name):
        return font_name
    windir = os.environ.get("WINDIR", r"C:\Windows")
    candidates = [
        os.path.join(windir, "Fonts", f"{font_name}.ttf"),
        os.path.join(windir, "Fonts", f"{font_name}.otf"),
        os.path.join(windir, "Fonts", f"{font_name.lower()}.ttf"),
        os.path.join(windir, "Fonts", f"{font_name.lower()}.otf"),
        os.path.join(windir, "Fonts", "arial.ttf"),
        os.path.join(windir, "Fonts", "calibri.ttf"),
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    ]
    for path in candidates:
        if os.path.isfile(path):
            return path
    return None


def load_font(font_name: str, size: int):
    if not PIL_AVAILABLE:
        return None
    font_path = find_font_file(font_name)
    if font_path:
        try:
            return ImageFont.truetype(font_path, size=size)
        except Exception:
            pass
    try:
        return ImageFont.truetype("arial.ttf", size=size)
    except Exception:
        return ImageFont.load_default()


def hex_to_rgb(value: str) -> Tuple[int, int, int]:
    if PIL_AVAILABLE:
        return ImageColor.getrgb(value or "#000000")
    return (0, 0, 0)


def draw_text_block(draw, box: Tuple[int, int, int, int], text: str, font, fill, align: str, spacing: int = 6):
    x, y, w, h = box
    words = text.split()
    lines: List[str] = []
    current = []
    for word in words or [""]:
        test = (" ".join(current + [word])).strip()
        bbox = draw.textbbox((0, 0), test, font=font)
        if current and (bbox[2] - bbox[0]) > w:
            lines.append(" ".join(current))
            current = [word]
        else:
            current.append(word)
    if current:
        lines.append(" ".join(current))

    cy = y
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=font)
        line_w = bbox[2] - bbox[0]
        line_h = bbox[3] - bbox[1]
        if align == "center":
            tx = x + max(0, (w - line_w) // 2)
        elif align == "right":
            tx = x + max(0, w - line_w)
        else:
            tx = x
        draw.text((tx, cy), line, font=font, fill=fill)
        cy += line_h + spacing
        if cy > y + h:
            break


def render_slide_to_image(slide: Slide, canvas_size: Tuple[int, int]) -> Optional[Image.Image]:
    if not PIL_AVAILABLE:
        return None
    width, height = canvas_size
    img = Image.new("RGB", (width, height), hex_to_rgb(PREVIEW_BG))
    if slide.background and os.path.exists(slide.background):
        try:
            bg = Image.open(slide.background).convert("RGB")
            bg = ImageOps.fit(bg, (width, height), centering=(slide.crop_x, slide.crop_y))
            img.paste(bg)
        except Exception:
            pass
    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    odraw = ImageDraw.Draw(overlay)
    r, g, b = hex_to_rgb(slide.style.overlay_color)
    alpha = max(0, min(255, int(slide.style.overlay_alpha)))
    odraw.rounded_rectangle((42, 42, width - 42, 175), radius=18, fill=(r, g, b, alpha))
    odraw.rounded_rectangle((42, height - 320, width - 42, height - 42), radius=20, fill=(r, g, b, alpha))
    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    draw = ImageDraw.Draw(img)
    title_font = load_font(slide.style.title_font, slide.style.title_size)
    body_font = load_font(slide.style.body_font, slide.style.body_size)
    fill = hex_to_rgb(slide.style.fg_color)
    draw_text_block(draw, (70, 72, width - 140, 88), slide.title, title_font, fill, slide.style.title_align, spacing=6)
    body = slide.body.strip() or slide.narration.strip() or "Slide text"
    draw_text_block(draw, (72, height - 292, width - 144, 220), body, body_font, fill, slide.style.body_align, spacing=8)
    return img


def wav_duration_seconds(path: Path) -> float:
    with wave.open(str(path), "rb") as wf:
        frames = wf.getnframes()
        rate = wf.getframerate()
        return 0.0 if rate <= 0 else frames / float(rate)


def concatenate_wavs(input_paths: List[Path], output_path: Path, gap_sec: float = 0.0) -> None:
    if not input_paths:
        raise ValueError("No input WAV files.")
    with wave.open(str(input_paths[0]), "rb") as wf:
        nchannels = wf.getnchannels()
        sampwidth = wf.getsampwidth()
        framerate = wf.getframerate()
        comptype = wf.getcomptype()
        compname = wf.getcompname()
    silence_frames = int(round(gap_sec * framerate))
    silence = b"\x00" * silence_frames * nchannels * sampwidth
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with wave.open(str(output_path), "wb") as out:
        out.setnchannels(nchannels)
        out.setsampwidth(sampwidth)
        out.setframerate(framerate)
        out.setcomptype(comptype, compname)
        for idx, path in enumerate(input_paths):
            with wave.open(str(path), "rb") as wf:
                out.writeframes(wf.readframes(wf.getnframes()))
            if gap_sec > 0 and idx < len(input_paths) - 1:
                out.writeframes(silence)


def detect_missing_packages() -> List[str]:
    missing = []
    if not PIL_AVAILABLE:
        missing.append("pillow")
    try:
        import torch  # noqa
    except Exception:
        missing.append("torch")
    try:
        import torchaudio  # noqa
    except Exception:
        missing.append("torchaudio")
    try:
        from chatterbox.tts_turbo import ChatterboxTurboTTS  # noqa
    except Exception:
        missing.append("chatterbox-tts")
    return sorted(set(missing))


def detect_optional_missing_packages() -> List[str]:
    missing = []
    if not FASTER_WHISPER_AVAILABLE:
        missing.append("faster-whisper")
    if not SOUNDDEVICE_AVAILABLE:
        missing.append("sounddevice")
    if not SOUNDFILE_AVAILABLE:
        missing.append("soundfile")
    return sorted(set(missing))


def pip_install(packages: List[str]) -> None:
    cmd = [sys.executable, "-m", "pip", "install", "--upgrade"] + packages
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or proc.stdout.strip() or "pip install failed")


def resolve_ffmpeg(configured_path: str) -> str:
    if configured_path and os.path.isfile(configured_path):
        return configured_path
    found = shutil.which("ffmpeg")
    if found:
        return found
    raise RuntimeError("ffmpeg not found. Put ffmpeg on PATH or choose ffmpeg.exe with Pick ffmpeg.")


def write_ffmpeg_concat_file(image_paths: List[Path], durations: List[float], output_path: Path) -> None:
    lines = []
    for image_path, duration in zip(image_paths, durations):
        safe = str(image_path.resolve()).replace("'", r"'\\''")
        lines.append(f"file '{safe}'")
        lines.append(f"duration {max(0.05, duration):.4f}")
    safe_last = str(image_paths[-1].resolve()).replace("'", r"'\\''")
    lines.append(f"file '{safe_last}'")
    output_path.write_text("\n".join(lines), encoding="utf-8")


def build_video_from_images_and_audio(image_paths: List[Path], durations: List[float], full_audio_wav: Path, output_mp4: Path, fps: int, configured_ffmpeg: str) -> None:
    ffmpeg = resolve_ffmpeg(configured_ffmpeg)
    concat_file = output_mp4.parent / "slides_concat.txt"
    write_ffmpeg_concat_file(image_paths, durations, concat_file)
    cmd = [
        ffmpeg,
        "-y",
        "-f", "concat",
        "-safe", "0",
        "-i", str(concat_file),
        "-i", str(full_audio_wav),
        "-vsync", "vfr",
        "-pix_fmt", "yuv420p",
        "-r", str(fps),
        "-c:v", "libx264",
        "-c:a", "aac",
        "-b:a", "192k",
        "-shortest",
        str(output_mp4),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or proc.stdout.strip() or "ffmpeg export failed")


class ChatterboxTurboEngine:
    def __init__(self, settings: AppSettings):
        self.settings = settings
        self._model = None
        self._ta = None

    def _load(self) -> str:
        if self._model is not None:
            return self._device
        try:
            import torch
            import torchaudio as ta
            from chatterbox.tts_turbo import ChatterboxTurboTTS
        except Exception as exc:
            raise RuntimeError("Chatterbox-Turbo is not ready. Run Setup first.\n\n" + str(exc))
        device = "cpu"
        if self.settings.use_cuda:
            try:
                if torch.cuda.is_available():
                    device = "cuda"
                elif hasattr(torch.backends, "mps") and torch.backends.mps.is_available():
                    device = "mps"
            except Exception:
                pass
        self._device = device
        self._ta = ta
        self._model = ChatterboxTurboTTS.from_pretrained(device=device)
        return device

    def warmup(self) -> str:
        device = self._load()
        return f"Chatterbox-Turbo ready on {device}"

    def synthesize(self, text: str, out_wav: Path) -> None:
        self._load()
        text = (text or "").strip() or "Hello. This is a test."
        kwargs = {
            "exaggeration": float(self.settings.chatterbox_exaggeration),
            "cfg_weight": float(self.settings.chatterbox_cfg_weight),
            "temperature": float(self.settings.chatterbox_temperature),
        }
        voice_sample = (self.settings.voice_sample_wav or "").strip()
        if voice_sample and os.path.isfile(voice_sample):
            wav = self._model.generate(text, audio_prompt_path=voice_sample, **kwargs)
        else:
            wav = self._model.generate(text, **kwargs)
        out_wav.parent.mkdir(parents=True, exist_ok=True)
        self._ta.save(str(out_wav), wav, self._model.sr)


class TranscriptionEngine:
    def __init__(self, settings: AppSettings):
        self.settings = settings
        self._model = None

    def transcribe(self, media_path: str) -> str:
        if not FASTER_WHISPER_AVAILABLE:
            raise RuntimeError("faster-whisper is not installed. Run Setup and include optional tools.")
        if not os.path.isfile(media_path):
            raise RuntimeError("Media file not found.")
        if self._model is None:
            compute_type = "float16" if self.settings.use_cuda else "int8"
            self._model = WhisperModel(self.settings.whisper_model or "small", device=self.settings.whisper_device or "auto", compute_type=compute_type)
        segments, _info = self._model.transcribe(media_path)
        return " ".join(seg.text.strip() for seg in segments if seg.text.strip()).strip()


class TimelineCanvas(tk.Canvas):
    def __init__(self, master, app: "TypeToVideoApp", **kwargs):
        super().__init__(master, bg=TIMELINE_BG, highlightthickness=0, **kwargs)
        self.app = app
        self.drag_mode = None
        self.drag_slide_id = None
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.orig_start = 0.0
        self.orig_duration = 0.0
        self.orig_layer = 0
        self.bind("<Button-1>", self.on_press)
        self.bind("<B1-Motion>", self.on_drag)
        self.bind("<ButtonRelease-1>", self.on_release)
        self.bind("<Double-Button-1>", self.on_double_click)

    def seconds_to_x(self, sec: float) -> float:
        return TIMELINE_HEADER_WIDTH + sec * PIXELS_PER_SECOND

    def layer_to_y(self, layer: int) -> float:
        return layer * TIMELINE_LAYER_HEIGHT

    def redraw(self) -> None:
        self.delete("all")
        project = self.app.project
        max_layer = max(3, project.max_layer() + 2)
        total_dur = max(20, math.ceil(project.total_duration()) + 4)
        width = int(self.seconds_to_x(total_dur)) + 220
        height = max_layer * TIMELINE_LAYER_HEIGHT
        self.config(scrollregion=(0, 0, width, height))
        for layer in range(max_layer):
            y1 = self.layer_to_y(layer)
            y2 = y1 + TIMELINE_LAYER_HEIGHT
            fill = "#111827" if layer % 2 == 0 else "#0b1220"
            self.create_rectangle(0, y1, width, y2, fill=fill, outline=TIMELINE_GRID)
            self.create_text(10, y1 + TIMELINE_LAYER_HEIGHT / 2, anchor="w", fill=TIMELINE_TEXT, text=f"L{layer}")
        for sec in range(total_dur + 1):
            x = self.seconds_to_x(sec)
            self.create_line(x, 0, x, height, fill=TIMELINE_GRID)
            self.create_text(x + 2, 10, anchor="nw", fill="#94A3B8", text=f"{sec}s", font=("Arial", 9))
        for slide in project.slides:
            self._draw_clip(slide)

    def _draw_clip(self, slide: Slide) -> None:
        x1 = self.seconds_to_x(slide.start)
        x2 = self.seconds_to_x(slide.end)
        y1 = self.layer_to_y(slide.layer) + 8
        y2 = y1 + TIMELINE_LAYER_HEIGHT - 16
        fill = CLIP_SELECTED if slide.slide_id == self.app.selected_slide_id else CLIP_FILL
        self.create_rectangle(x1, y1, x2, y2, fill=fill, outline="#D1D5DB", width=1, tags=("clip", f"slide:{slide.slide_id}"))
        self.create_rectangle(x1, y1, x1 + 7, y2, fill="#1D4ED8", outline="", tags=("left_handle", f"slide:{slide.slide_id}"))
        self.create_rectangle(x2 - 7, y1, x2, y2, fill="#1D4ED8", outline="", tags=("right_handle", f"slide:{slide.slide_id}"))
        label = f"{slide.title} [{slide.duration:.1f}s]"
        self.create_text(x1 + 10, y1 + 8, anchor="nw", fill=CLIP_TEXT, text=label, width=max(50, x2 - x1 - 18), font=("Arial", 10), tags=("clip_text", f"slide:{slide.slide_id}"))

    def _find_hit(self, event) -> Tuple[Optional[str], Optional[str]]:
        item = self.find_closest(event.x, event.y)
        if not item:
            return None, None
        tags = self.gettags(item[0])
        slide_id = None
        role = None
        for tag in tags:
            if tag.startswith("slide:"):
                slide_id = tag.split(":", 1)[1]
            elif tag in {"clip", "clip_text", "left_handle", "right_handle"}:
                role = tag
        return slide_id, role

    def on_press(self, event) -> None:
        slide_id, role = self._find_hit(event)
        if not slide_id:
            return
        slide = self.app.get_slide(slide_id)
        if slide is None:
            return
        self.app.select_slide(slide_id)
        self.drag_slide_id = slide_id
        self.drag_start_x = event.x
        self.drag_start_y = event.y
        self.orig_start = slide.start
        self.orig_duration = slide.duration
        self.orig_layer = slide.layer
        if role == "left_handle":
            self.drag_mode = "trim_left"
        elif role == "right_handle":
            self.drag_mode = "trim_right"
        else:
            self.drag_mode = "move"

    def on_drag(self, event) -> None:
        if not self.drag_slide_id or not self.drag_mode:
            return
        slide = self.app.get_slide(self.drag_slide_id)
        if slide is None:
            return
        dx_sec = (event.x - self.drag_start_x) / PIXELS_PER_SECOND
        dy = event.y - self.drag_start_y
        if self.drag_mode == "move":
            slide.start = max(0.0, round(self.orig_start + dx_sec, 2))
            slide.layer = max(0, self.orig_layer + int(round(dy / TIMELINE_LAYER_HEIGHT)))
        elif self.drag_mode == "trim_left":
            new_start = max(0.0, round(self.orig_start + dx_sec, 2))
            max_start = self.orig_start + self.orig_duration - TIMELINE_MIN_DURATION
            new_start = min(new_start, max_start)
            slide.duration = round((self.orig_start + self.orig_duration) - new_start, 2)
            slide.start = new_start
        elif self.drag_mode == "trim_right":
            slide.duration = max(TIMELINE_MIN_DURATION, round(self.orig_duration + dx_sec, 2))
        self.app.refresh_slide_editor()
        self.app.refresh_slide_list()
        self.app.sync_project_source()
        self.redraw()
        self.app.render_preview()

    def on_release(self, event) -> None:
        self.drag_mode = None
        self.drag_slide_id = None
        self.app.mark_dirty()

    def on_double_click(self, event) -> None:
        slide_id, _role = self._find_hit(event)
        if slide_id:
            self.app.select_slide(slide_id)
            self.app.body_text.focus_set()


class TypeToVideoApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.current_path: Optional[Path] = None
        self.is_dirty = False
        self.selected_slide_id: Optional[str] = None
        self.preview_image_cache = None
        self._drag_crop = False
        self._chatterbox_engine: Optional[ChatterboxTurboEngine] = None
        self.project = self._default_project()
        self._build_ui()
        self.select_slide(self.project.slides[0].slide_id if self.project.slides else None)
        self.refresh_script_text()
        self.refresh_settings_ui()
        self.sync_project_source()
        self.timeline.redraw()
        self.render_preview()
        self.update_title()

    def _default_project(self) -> Project:
        raw = "Welcome. Type your message here. This app turns plain text into a narrated video. Paste a script, click Make Slides, and export when you are ready."
        settings = AppSettings()
        slides = slides_from_raw_text(raw, settings)
        return Project(name="Untitled Project", raw_script=raw, settings=settings, slides=slides)

    def _build_menu(self) -> None:
        menu = tk.Menu(self.root)
        file_menu = tk.Menu(menu, tearoff=False)
        file_menu.add_command(label="New", command=self.new_project)
        file_menu.add_command(label="Open", command=self.open_project)
        file_menu.add_command(label="Save", command=self.save_project)
        file_menu.add_command(label="Save As", command=self.save_project_as)
        file_menu.add_separator()
        file_menu.add_command(label="Export MP4", command=self.export_mp4_dialog)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menu.add_cascade(label="File", menu=file_menu)

        tools_menu = tk.Menu(menu, tearoff=False)
        tools_menu.add_command(label="Setup", command=self.setup_dialog)
        tools_menu.add_command(label="Warm Up Voice", command=self.warmup_chatterbox)
        tools_menu.add_command(label="Make Slides", command=self.build_slides_from_script)
        tools_menu.add_command(label="Import Audio", command=self.import_audio_and_transcribe)
        tools_menu.add_command(label="Record Audio", command=self.record_audio_and_transcribe)
        tools_menu.add_command(label="Rewrite With Ollama", command=self.ai_rewrite_raw_script)
        tools_menu.add_command(label="Rebuild From Project Code", command=self.rebuild_from_project_source)
        menu.add_cascade(label="Tools", menu=tools_menu)
        self.root.config(menu=menu)

    def _build_ui(self) -> None:
        self.root.geometry("1780x1040")
        self.root.minsize(1450, 860)
        self._build_menu()

        topbar = ttk.Frame(self.root)
        topbar.pack(fill="x", padx=8, pady=8)
        ttk.Button(topbar, text="Setup", command=self.setup_dialog).pack(side="left")
        ttk.Button(topbar, text="Pick ffmpeg", command=self.choose_ffmpeg).pack(side="left", padx=4)
        ttk.Button(topbar, text="Warm Up Voice", command=self.warmup_chatterbox).pack(side="left", padx=4)
        ttk.Button(topbar, text="Import Audio", command=self.import_audio_and_transcribe).pack(side="left", padx=4)
        ttk.Button(topbar, text="Record Audio", command=self.record_audio_and_transcribe).pack(side="left", padx=4)
        ttk.Button(topbar, text="Rewrite", command=self.ai_rewrite_raw_script).pack(side="left", padx=4)
        ttk.Button(topbar, text="Make Slides", command=self.build_slides_from_script).pack(side="left", padx=4)
        ttk.Button(topbar, text="Test Voice", command=self.test_voice_dialog).pack(side="left", padx=4)
        ttk.Button(topbar, text="Export MP4", command=self.export_mp4_dialog).pack(side="left", padx=10)

        ttk.Label(topbar, text="Voice Sample").pack(side="left", padx=(18, 6))
        self.voice_sample_var = tk.StringVar()
        ttk.Entry(topbar, textvariable=self.voice_sample_var, width=44).pack(side="left", fill="x", expand=True)
        ttk.Button(topbar, text="Browse", command=self.choose_voice_sample).pack(side="left", padx=4)

        main = ttk.Panedwindow(self.root, orient=tk.HORIZONTAL)
        main.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        left = ttk.Frame(main)
        right = ttk.Frame(main)
        main.add(left, weight=3)
        main.add(right, weight=2)
        self._build_left(left)
        self._build_right(right)

        timeline_group = ttk.LabelFrame(self.root, text="Timeline")
        timeline_group.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.timeline = TimelineCanvas(timeline_group, self, height=270)
        x_scroll = ttk.Scrollbar(timeline_group, orient="horizontal", command=self.timeline.xview)
        y_scroll = ttk.Scrollbar(timeline_group, orient="vertical", command=self.timeline.yview)
        self.timeline.configure(xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)
        self.timeline.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        timeline_group.rowconfigure(0, weight=1)
        timeline_group.columnconfigure(0, weight=1)

        status = ttk.Frame(self.root)
        status.pack(fill="x", side="bottom")
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(status, textvariable=self.status_var).pack(anchor="w", padx=8, pady=4)

    def _build_left(self, parent) -> None:
        group = ttk.LabelFrame(parent, text="1. Type Or Paste Your Script")
        group.pack(fill="both", expand=True)
        helper = ttk.Frame(group)
        helper.pack(fill="x", padx=8, pady=8)
        self.max_chars_var = tk.StringVar()
        self.min_chars_var = tk.StringVar()
        self.background_folder_var = tk.StringVar()
        ttk.Label(helper, text="Max chars/slide").pack(side="left")
        ttk.Entry(helper, textvariable=self.max_chars_var, width=8).pack(side="left", padx=(4, 12))
        ttk.Label(helper, text="Min chars/slide").pack(side="left")
        ttk.Entry(helper, textvariable=self.min_chars_var, width=8).pack(side="left", padx=(4, 12))
        ttk.Label(helper, text="Background folder").pack(side="left")
        ttk.Entry(helper, textvariable=self.background_folder_var, width=30).pack(side="left", padx=(4, 4))
        ttk.Button(helper, text="Browse", command=self.choose_background_folder).pack(side="left", padx=(0, 8))
        self.script_text = tk.Text(group, wrap="word", font=("Consolas", 12))
        self.script_text.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def _build_right(self, parent) -> None:
        preview_group = ttk.LabelFrame(parent, text="2. Preview")
        preview_group.pack(fill="both", expand=True, pady=(0, 8))
        self.preview_canvas = tk.Canvas(preview_group, bg=PREVIEW_BG, highlightthickness=1, highlightbackground="#334155")
        self.preview_canvas.pack(fill="both", expand=True, padx=8, pady=8)
        self.preview_canvas.bind("<ButtonPress-1>", self.on_preview_crop_press)
        self.preview_canvas.bind("<B1-Motion>", self.on_preview_crop_drag)
        self.preview_canvas.bind("<ButtonRelease-1>", self.on_preview_crop_release)

        bottom = ttk.Panedwindow(parent, orient=tk.VERTICAL)
        bottom.pack(fill="both", expand=True)
        slides_group = ttk.LabelFrame(bottom, text="3. Slides")
        edit_group = ttk.LabelFrame(bottom, text="4. Edit Selected Slide")
        bottom.add(slides_group, weight=1)
        bottom.add(edit_group, weight=1)

        self.slide_list = tk.Listbox(slides_group, exportselection=False)
        self.slide_list.pack(fill="both", expand=True, padx=8, pady=8)
        self.slide_list.bind("<<ListboxSelect>>", self.on_slide_list_select)

        toolbar = ttk.Frame(edit_group)
        toolbar.pack(fill="x", padx=8, pady=(8, 4))
        ttk.Button(toolbar, text="Add", command=self.add_slide).pack(side="left")
        ttk.Button(toolbar, text="Duplicate", command=self.duplicate_slide).pack(side="left", padx=4)
        ttk.Button(toolbar, text="Delete", command=self.delete_slide).pack(side="left", padx=4)
        ttk.Button(toolbar, text="Apply", command=self.apply_selected_slide).pack(side="left", padx=10)
        ttk.Button(toolbar, text="Refresh Code", command=self.sync_project_source).pack(side="left", padx=4)

        form = ttk.Frame(edit_group)
        form.pack(fill="x", padx=8, pady=(0, 6))
        self.title_var = tk.StringVar()
        self.duration_var = tk.StringVar()
        self.background_var = tk.StringVar()
        ttk.Label(form, text="Title").grid(row=0, column=0, sticky="w", pady=3)
        self.title_entry = ttk.Entry(form, textvariable=self.title_var)
        self.title_entry.grid(row=0, column=1, sticky="ew", pady=3)
        ttk.Label(form, text="Duration").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(form, textvariable=self.duration_var).grid(row=1, column=1, sticky="ew", pady=3)
        ttk.Label(form, text="Background").grid(row=2, column=0, sticky="w", pady=3)
        bg_row = ttk.Frame(form)
        bg_row.grid(row=2, column=1, sticky="ew", pady=3)
        ttk.Entry(bg_row, textvariable=self.background_var).pack(side="left", fill="x", expand=True)
        ttk.Button(bg_row, text="Browse", command=self.choose_current_slide_background).pack(side="left", padx=(4, 0))
        form.columnconfigure(1, weight=1)

        ttk.Label(edit_group, text="Body").pack(anchor="w", padx=8)
        self.body_text = tk.Text(edit_group, height=7, wrap="word", font=BODY_FONT)
        self.body_text.pack(fill="both", expand=True, padx=8, pady=(0, 6))
        ttk.Label(edit_group, text="Narration").pack(anchor="w", padx=8)
        self.narration_text = tk.Text(edit_group, height=7, wrap="word", font=BODY_FONT)
        self.narration_text.pack(fill="both", expand=True, padx=8, pady=(0, 6))

        code_group = ttk.LabelFrame(parent, text="Project Code")
        code_group.pack(fill="both", expand=True, pady=(8, 0))
        self.project_source = tk.Text(code_group, height=10, wrap="none", font=("Consolas", 10))
        self.project_source.pack(fill="both", expand=True, padx=8, pady=8)

    def set_status(self, text: str) -> None:
        self.status_var.set(text)
        self.root.update_idletasks()

    def mark_dirty(self) -> None:
        self.is_dirty = True
        self.update_title()

    def mark_clean(self) -> None:
        self.is_dirty = False
        self.update_title()

    def update_title(self) -> None:
        star = "*" if self.is_dirty else ""
        path = str(self.current_path) if self.current_path else "Untitled"
        self.root.title(f"{APP_TITLE} - {path}{star}")

    def get_slide(self, slide_id: Optional[str]) -> Optional[Slide]:
        if not slide_id:
            return None
        for slide in self.project.slides:
            if slide.slide_id == slide_id:
                return slide
        return None

    def refresh_script_text(self) -> None:
        self.script_text.delete("1.0", tk.END)
        self.script_text.insert("1.0", self.project.raw_script)

    def refresh_settings_ui(self) -> None:
        s = self.project.settings
        self.max_chars_var.set(str(s.max_chars_per_slide))
        self.min_chars_var.set(str(s.min_chars_per_slide))
        self.voice_sample_var.set(s.voice_sample_wav)
        self.background_folder_var.set(s.background_folder)

    def refresh_slide_list(self) -> None:
        self.project.normalize_ids()
        self.project.sort_slides()
        self.slide_list.delete(0, tk.END)
        for slide in self.project.slides:
            self.slide_list.insert(tk.END, f"[{slide.layer}] {slide.start:>5.1f}s  {slide.title}")
        if self.selected_slide_id:
            for i, slide in enumerate(self.project.slides):
                if slide.slide_id == self.selected_slide_id:
                    self.slide_list.selection_clear(0, tk.END)
                    self.slide_list.selection_set(i)
                    self.slide_list.see(i)
                    break

    def select_slide(self, slide_id: Optional[str]) -> None:
        self.selected_slide_id = slide_id
        self.refresh_slide_list()
        self.refresh_slide_editor()
        self.timeline.redraw()
        self.render_preview()

    def refresh_slide_editor(self) -> None:
        slide = self.get_slide(self.selected_slide_id)
        if slide is None:
            self.title_var.set("")
            self.duration_var.set("")
            self.background_var.set("")
            self.body_text.delete("1.0", tk.END)
            self.narration_text.delete("1.0", tk.END)
            return
        self.title_var.set(slide.title)
        self.duration_var.set(f"{slide.duration:.2f}")
        self.background_var.set(slide.background)
        self.body_text.delete("1.0", tk.END)
        self.body_text.insert("1.0", slide.body)
        self.narration_text.delete("1.0", tk.END)
        self.narration_text.insert("1.0", slide.narration)

    def sync_project_source(self) -> None:
        self.project_source.delete("1.0", tk.END)
        self.project_source.insert("1.0", plain_project_source(self.project))

    def apply_settings_from_ui(self) -> None:
        s = self.project.settings
        try:
            s.max_chars_per_slide = max(80, int(self.max_chars_var.get().strip() or 260))
            s.min_chars_per_slide = max(20, int(self.min_chars_var.get().strip() or 70))
        except ValueError:
            s.max_chars_per_slide = 260
            s.min_chars_per_slide = 70
        s.voice_sample_wav = self.voice_sample_var.get().strip()
        s.background_folder = self.background_folder_var.get().strip()
        self.mark_dirty()

    def _sync_ui_to_project(self) -> None:
        self.project.raw_script = self.script_text.get("1.0", tk.END).strip()
        self.apply_settings_from_ui()
        slide = self.get_slide(self.selected_slide_id)
        if slide:
            try:
                slide.title = self.title_var.get().strip() or "Untitled Slide"
                slide.duration = max(TIMELINE_MIN_DURATION, round(float(self.duration_var.get().strip() or 5), 2))
                slide.background = self.background_var.get().strip()
            except ValueError:
                pass
            slide.body = self.body_text.get("1.0", tk.END).strip()
            slide.narration = self.narration_text.get("1.0", tk.END).strip()
        self.project.normalize_ids()
        self.project.sort_slides()
        self.sync_project_source()

    def confirm_discard_changes(self) -> bool:
        if not self.is_dirty:
            return True
        return messagebox.askyesno("Unsaved changes", "Discard unsaved changes?")

    def new_project(self) -> None:
        if not self.confirm_discard_changes():
            return
        self.project = self._default_project()
        self.current_path = None
        self.selected_slide_id = self.project.slides[0].slide_id if self.project.slides else None
        self.refresh_script_text()
        self.refresh_settings_ui()
        self.refresh_slide_list()
        self.refresh_slide_editor()
        self.sync_project_source()
        self.timeline.redraw()
        self.render_preview()
        self.mark_clean()
        self.set_status("New project created")

    def open_project(self) -> None:
        if not self.confirm_discard_changes():
            return
        path = filedialog.askopenfilename(title="Open project", filetypes=[("Type2Video Project", f"*{PROJECT_EXTENSION}"), ("JSON", "*.json"), ("All files", "*.*")])
        if not path:
            return
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
            self.project = project_from_dict(data)
            self.current_path = Path(path)
            self.selected_slide_id = self.project.slides[0].slide_id if self.project.slides else None
            self.refresh_script_text()
            self.refresh_settings_ui()
            self.refresh_slide_list()
            self.refresh_slide_editor()
            self.sync_project_source()
            self.timeline.redraw()
            self.render_preview()
            self.mark_clean()
            self.set_status(f"Opened: {path}")
        except Exception as exc:
            messagebox.showerror("Open failed", str(exc))

    def save_project(self) -> None:
        if self.current_path is None:
            self.save_project_as()
            return
        self._sync_ui_to_project()
        try:
            self.current_path.write_text(json.dumps(project_to_dict(self.project), indent=2), encoding="utf-8")
            self.mark_clean()
            self.set_status(f"Saved: {self.current_path}")
        except Exception as exc:
            messagebox.showerror("Save failed", str(exc))

    def save_project_as(self) -> None:
        path = filedialog.asksaveasfilename(title="Save project", defaultextension=PROJECT_EXTENSION, filetypes=[("Type2Video Project", f"*{PROJECT_EXTENSION}"), ("JSON", "*.json"), ("All files", "*.*")])
        if not path:
            return
        self.current_path = Path(path)
        self.save_project()

    def choose_voice_sample(self) -> None:
        path = filedialog.askopenfilename(title="Choose voice sample", filetypes=[("Audio", "*.wav *.mp3 *.m4a *.flac"), ("All files", "*.*")])
        if path:
            self.voice_sample_var.set(path)
            self.apply_settings_from_ui()
            self.set_status("Voice sample selected")

    def choose_background_folder(self) -> None:
        path = filedialog.askdirectory(title="Choose background image folder")
        if path:
            self.background_folder_var.set(path)
            self.apply_settings_from_ui()
            self.set_status("Background folder selected")

    def choose_current_slide_background(self) -> None:
        slide = self.get_slide(self.selected_slide_id)
        if slide is None:
            return
        path = filedialog.askopenfilename(title="Choose slide background", filetypes=[("Images", "*.png *.jpg *.jpeg *.webp *.bmp"), ("All files", "*.*")])
        if path:
            slide.background = path
            self.background_var.set(path)
            self.sync_project_source()
            self.render_preview()
            self.mark_dirty()

    def choose_ffmpeg(self) -> None:
        path = filedialog.askopenfilename(title="Choose ffmpeg executable", filetypes=[("Executable", "*.exe"), ("All files", "*.*")])
        if path:
            self.project.settings.ffmpeg_path = path
            self.mark_dirty()
            self.set_status("ffmpeg selected")

    def build_slides_from_script(self) -> None:
        self.apply_settings_from_ui()
        raw = self.script_text.get("1.0", tk.END).strip()
        if not raw:
            messagebox.showinfo("No script", "Type or paste a script first.")
            return
        template = self.get_slide(self.selected_slide_id).style if self.get_slide(self.selected_slide_id) else TextStyle()
        slides = slides_from_raw_text(raw, self.project.settings, template)
        if not slides:
            messagebox.showerror("Build failed", "Could not build slides from the script.")
            return
        folder = (self.project.settings.background_folder or "").strip()
        bg_files = []
        if folder and os.path.isdir(folder):
            for name in sorted(os.listdir(folder)):
                if name.lower().endswith((".png", ".jpg", ".jpeg", ".webp", ".bmp")):
                    bg_files.append(os.path.join(folder, name))
        for i, slide in enumerate(slides):
            if bg_files:
                slide.background = bg_files[i % len(bg_files)]
        self.project.raw_script = raw
        self.project.slides = slides
        self.project.normalize_ids()
        self.project.sort_slides()
        self.select_slide(self.project.slides[0].slide_id if self.project.slides else None)
        self.sync_project_source()
        self.mark_dirty()
        self.set_status(f"Built {len(slides)} slides")

    def add_slide(self) -> None:
        self._sync_ui_to_project()
        idx = len(self.project.slides) + 1
        template_style = self.project.slides[-1].style if self.project.slides else TextStyle()
        slide = Slide(slide_id=f"slide_{idx:03d}", title=f"Slide {idx}", start=round(self.project.total_duration(), 2), duration=5.0, layer=0, body="Add text here.", narration="Add narration here.", style=TextStyle(**asdict(template_style)))
        self.project.slides.append(slide)
        self.project.sort_slides()
        self.select_slide(slide.slide_id)
        self.sync_project_source()
        self.mark_dirty()

    def duplicate_slide(self) -> None:
        slide = self.get_slide(self.selected_slide_id)
        if slide is None:
            return
        new_slide = slide.clone()
        new_slide.slide_id = ""
        new_slide.title = f"{slide.title} Copy"
        new_slide.start = round(slide.end + 0.25, 2)
        self.project.slides.append(new_slide)
        self.project.normalize_ids()
        self.project.sort_slides()
        self.select_slide(new_slide.slide_id)
        self.sync_project_source()
        self.mark_dirty()

    def delete_slide(self) -> None:
        slide = self.get_slide(self.selected_slide_id)
        if slide is None:
            return
        if not messagebox.askyesno("Delete slide", f"Delete '{slide.title}'?"):
            return
        self.project.slides = [s for s in self.project.slides if s.slide_id != slide.slide_id]
        self.project.normalize_ids()
        self.project.sort_slides()
        self.select_slide(self.project.slides[0].slide_id if self.project.slides else None)
        self.sync_project_source()
        self.mark_dirty()

    def apply_selected_slide(self) -> None:
        slide = self.get_slide(self.selected_slide_id)
        if slide is None:
            return
        try:
            slide.title = self.title_var.get().strip() or "Untitled Slide"
            slide.duration = max(TIMELINE_MIN_DURATION, round(float(self.duration_var.get().strip() or 5), 2))
            slide.background = self.background_var.get().strip()
        except ValueError:
            messagebox.showerror("Invalid value", "Duration must be a number.")
            return
        slide.body = self.body_text.get("1.0", tk.END).strip()
        slide.narration = self.narration_text.get("1.0", tk.END).strip()
        self.project.sort_slides()
        self.refresh_slide_list()
        self.sync_project_source()
        self.timeline.redraw()
        self.render_preview()
        self.mark_dirty()
        self.set_status("Slide updated")

    def on_slide_list_select(self, event=None) -> None:
        selection = self.slide_list.curselection()
        if not selection:
            return
        idx = selection[0]
        if 0 <= idx < len(self.project.slides):
            self.select_slide(self.project.slides[idx].slide_id)

    def render_preview(self) -> None:
        self.preview_canvas.delete("all")
        slide = self.get_slide(self.selected_slide_id)
        canvas_w = max(1, self.preview_canvas.winfo_width())
        canvas_h = max(1, self.preview_canvas.winfo_height())
        self.preview_canvas.create_rectangle(0, 0, canvas_w, canvas_h, fill=PREVIEW_BG, outline="")
        if slide is None:
            self.preview_canvas.create_text(canvas_w / 2, canvas_h / 2, text="No slide selected", fill="#CBD5E1", font=TITLE_FONT)
            return
        if PIL_AVAILABLE:
            img = render_slide_to_image(slide, (canvas_w, canvas_h))
            if img is not None:
                self.preview_image_cache = ImageTk.PhotoImage(img)
                self.preview_canvas.create_image(0, 0, image=self.preview_image_cache, anchor="nw")
        else:
            self.preview_canvas.create_text(canvas_w / 2, canvas_h / 2, text="Install pillow for preview", fill="#CBD5E1", font=TITLE_FONT)
        self.preview_canvas.create_text(canvas_w - 12, 10, anchor="ne", text=f"{slide.start:.2f}s • {slide.duration:.2f}s • L{slide.layer}", fill="#CBD5E1", font=("Arial", 10))

    def on_preview_crop_press(self, event) -> None:
        if self.get_slide(self.selected_slide_id) is None:
            return
        self._drag_crop = True
        self.update_crop_from_preview(event.x, event.y)

    def on_preview_crop_drag(self, event) -> None:
        if self._drag_crop:
            self.update_crop_from_preview(event.x, event.y)

    def on_preview_crop_release(self, event) -> None:
        if self._drag_crop:
            self._drag_crop = False
            self.update_crop_from_preview(event.x, event.y)
            self.mark_dirty()

    def update_crop_from_preview(self, x: float, y: float) -> None:
        slide = self.get_slide(self.selected_slide_id)
        if slide is None:
            return
        w = max(1, self.preview_canvas.winfo_width())
        h = max(1, self.preview_canvas.winfo_height())
        slide.crop_x = min(1.0, max(0.0, x / w))
        slide.crop_y = min(1.0, max(0.0, y / h))
        self.render_preview()
        self.sync_project_source()

    def setup_dialog(self) -> None:
        missing_core = detect_missing_packages()
        missing_optional = detect_optional_missing_packages()
        lines = [
            f"Python: {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
            f"OS: {platform.system()} {platform.release()}",
            "",
            f"Recommended for Chatterbox: Python {RECOMMENDED_PYTHON}",
            "",
            "Core missing packages:",
            "  " + (", ".join(missing_core) if missing_core else "none"),
            "",
            "Optional missing packages:",
            "  " + (", ".join(missing_optional) if missing_optional else "none"),
            "",
            "ffmpeg on PATH:",
            "  yes" if shutil.which("ffmpeg") else "  no",
            "",
            "Yes = install core + optional",
            "No = install core only",
            "Cancel = close",
        ]
        choice = messagebox.askyesnocancel("Setup", "\n".join(lines))
        if choice is None:
            return
        packages = list(missing_core)
        if choice is True:
            packages.extend(missing_optional)
        packages = sorted(set(packages))
        if not packages:
            messagebox.showinfo("Setup", "Nothing obvious is missing.")
            return
        self.run_in_thread(lambda: pip_install(packages), "Installing packages...")

    def get_chatterbox_engine(self) -> ChatterboxTurboEngine:
        if self._chatterbox_engine is None:
            self._chatterbox_engine = ChatterboxTurboEngine(self.project.settings)
        else:
            self._chatterbox_engine.settings = self.project.settings
        return self._chatterbox_engine

    def warmup_chatterbox(self) -> None:
        self.apply_settings_from_ui()
        self.run_in_thread(self._warmup_chatterbox_worker, "Downloading / warming up Chatterbox-Turbo...")

    def _warmup_chatterbox_worker(self) -> None:
        msg = self.get_chatterbox_engine().warmup()
        self.root.after(0, lambda: messagebox.showinfo("Voice Ready", msg))

    def import_audio_and_transcribe(self) -> None:
        path = filedialog.askopenfilename(title="Select audio or video", filetypes=[("Media", "*.wav *.mp3 *.m4a *.flac *.mp4 *.mov *.mkv"), ("All files", "*.*")])
        if not path:
            return
        self.apply_settings_from_ui()
        self.run_in_thread(lambda: self._transcribe_worker(path), "Transcribing media...")

    def _transcribe_worker(self, path: str) -> None:
        transcript = TranscriptionEngine(self.project.settings).transcribe(path)
        self.root.after(0, lambda: self._append_transcript(transcript, os.path.basename(path)))

    def record_audio_and_transcribe(self) -> None:
        if not (SOUNDDEVICE_AVAILABLE and SOUNDFILE_AVAILABLE):
            messagebox.showerror("Recording unavailable", "Install sounddevice and soundfile from Setup first.")
            return
        seconds = simpledialog.askinteger("Record", "How many seconds?", initialvalue=20, minvalue=1, maxvalue=600)
        if not seconds:
            return
        out_dir = Path.home() / "Type2Video_Recordings"
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"recording_{int(time.time())}.wav"
        self.run_in_thread(lambda: self._record_and_transcribe_worker(seconds, out_path), "Recording and transcribing...")

    def _record_and_transcribe_worker(self, seconds: int, out_path: Path) -> None:
        fs = 16000
        audio = sd.rec(int(seconds * fs), samplerate=fs, channels=1, dtype="float32")
        sd.wait()
        sf.write(str(out_path), audio, fs)
        transcript = TranscriptionEngine(self.project.settings).transcribe(str(out_path))
        self.root.after(0, lambda: self._append_transcript(transcript, out_path.name))

    def _append_transcript(self, transcript: str, label: str) -> None:
        if not transcript.strip():
            messagebox.showwarning("No transcript", "Transcription returned no text.")
            return
        current = self.script_text.get("1.0", tk.END).strip()
        combined = (current + "\n\n" + transcript.strip()).strip() if current else transcript.strip()
        self.project.raw_script = combined
        self.refresh_script_text()
        self.mark_dirty()
        self.set_status(f"Added transcript from {label}")

    def ai_rewrite_raw_script(self) -> None:
        raw = self.script_text.get("1.0", tk.END).strip()
        if not raw:
            messagebox.showinfo("No text", "Type or transcribe some text first.")
            return
        prompt = "Rewrite the following into clean, natural spoken video script text. Preserve meaning. Keep it easy to segment into slides. Avoid markdown bullets unless necessary.\n\n" + raw
        self.run_in_thread(lambda: self._ollama_worker(prompt), "Rewriting script with local AI...")

    def _ollama_worker(self, prompt: str) -> None:
        payload = json.dumps({"model": self.project.settings.ollama_model, "prompt": prompt, "stream": False}).encode("utf-8")
        req = urllib.request.Request(self.project.settings.ollama_url, data=payload, headers={"Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=300) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except urllib.error.URLError as exc:
            self.root.after(0, lambda: messagebox.showerror("Ollama error", f"Could not reach local Ollama.\n\n{exc}"))
            return
        text = (data.get("response") or "").strip()
        self.root.after(0, lambda: self._replace_script_with_ai_text(text))

    def _replace_script_with_ai_text(self, text: str) -> None:
        if not text:
            messagebox.showwarning("No output", "Local AI returned no text.")
            return
        self.project.raw_script = text
        self.refresh_script_text()
        self.mark_dirty()
        self.set_status("Script updated from local AI")

    def rebuild_from_project_source(self) -> None:
        try:
            self.project = parse_plain_project_source(self.project_source.get("1.0", tk.END), self.project.settings, self.script_text.get("1.0", tk.END).strip())
            self.selected_slide_id = self.project.slides[0].slide_id if self.project.slides else None
            self.refresh_slide_list()
            self.refresh_slide_editor()
            self.timeline.redraw()
            self.render_preview()
            self.mark_dirty()
            self.set_status("Rebuilt slides from project code")
        except Exception as exc:
            messagebox.showerror("Parse error", str(exc))

    def test_voice_dialog(self) -> None:
        self.apply_settings_from_ui()
        text = self.narration_text.get("1.0", tk.END).strip() if self.get_slide(self.selected_slide_id) else ""
        if not text:
            text = self.script_text.get("1.0", tk.END).strip()
        if not text:
            messagebox.showinfo("No text", "Type some text first.")
            return
        test_text = textwrap.shorten(text.replace("\n", " "), width=220, placeholder=" ...")
        out_path = filedialog.asksaveasfilename(title="Save test voice WAV", defaultextension=".wav", filetypes=[("WAV", "*.wav"), ("All files", "*.*")])
        if not out_path:
            return
        self.run_in_thread(lambda: self._test_voice_worker(test_text, Path(out_path)), "Generating test voice...")

    def _test_voice_worker(self, text: str, out_path: Path) -> None:
        self.get_chatterbox_engine().synthesize(text, out_path)
        self.root.after(0, lambda: messagebox.showinfo("Voice Saved", f"Saved test voice to:\n\n{out_path}"))

    def export_mp4_dialog(self) -> None:
        if not self.project.slides:
            messagebox.showerror("No slides", "There are no slides to export.")
            return
        self._sync_ui_to_project()
        path = filedialog.asksaveasfilename(title="Export MP4", defaultextension=".mp4", filetypes=[("MP4", "*.mp4"), ("All files", "*.*")])
        if not path:
            return
        self.run_in_thread(lambda: self._export_mp4_worker(Path(path)), "Exporting MP4...")

    def _export_mp4_worker(self, output_mp4: Path) -> None:
        if not PIL_AVAILABLE:
            raise RuntimeError("Pillow is required for rendering. Install it from Setup.")
        engine = self.get_chatterbox_engine()
        temp_dir = output_mp4.parent / f"{output_mp4.stem}_render"
        images_dir = temp_dir / "images"
        audio_dir = temp_dir / "audio"
        images_dir.mkdir(parents=True, exist_ok=True)
        audio_dir.mkdir(parents=True, exist_ok=True)

        image_paths: List[Path] = []
        wav_paths: List[Path] = []
        durations: List[float] = []

        for idx, slide in enumerate(self.project.slides, 1):
            self.root.after(0, lambda i=idx: self.set_status(f"Rendering slide {i}/{len(self.project.slides)}"))
            img = render_slide_to_image(slide, (self.project.settings.canvas_width, self.project.settings.canvas_height))
            if img is None:
                raise RuntimeError("Failed to render slide image.")
            img_path = images_dir / f"slide_{idx:03d}.png"
            img.save(img_path)
            image_paths.append(img_path)
            narr_text = (slide.narration or slide.body or slide.title or f"Slide {idx}").strip()
            wav_path = audio_dir / f"slide_{idx:03d}.wav"
            self.root.after(0, lambda i=idx: self.set_status(f"Generating voice {i}/{len(self.project.slides)}"))
            engine.synthesize(narr_text, wav_path)
            wav_paths.append(wav_path)
            actual_dur = wav_duration_seconds(wav_path) + self.project.settings.gap_after_slide_sec
            slide.duration = max(slide.duration, round(actual_dur, 2))
            durations.append(slide.duration)

        concatenate_wavs(wav_paths, temp_dir / "full_audio.wav", gap_sec=self.project.settings.gap_after_slide_sec)
        t = 0.0
        for slide in self.project.slides:
            slide.start = round(t, 2)
            t += slide.duration
        build_video_from_images_and_audio(image_paths, durations, temp_dir / "full_audio.wav", output_mp4, self.project.settings.fps, self.project.settings.ffmpeg_path)
        self.root.after(0, lambda: messagebox.showinfo("Export Complete", f"Saved video to:\n\n{output_mp4}"))

    def run_in_thread(self, fn, status_text: str) -> None:
        self.set_status(status_text)
        def runner():
            try:
                fn()
            except Exception as exc:
                self.root.after(0, lambda: messagebox.showerror("Operation failed", str(exc)))
            finally:
                self.root.after(0, lambda: self.set_status("Ready"))
        threading.Thread(target=runner, daemon=True).start()


def main() -> None:
    root = tk.Tk()
    try:
        style = ttk.Style(root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass
    TypeToVideoApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
