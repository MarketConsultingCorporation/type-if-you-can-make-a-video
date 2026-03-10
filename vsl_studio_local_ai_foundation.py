#!/usr/bin/env python3
"""
VSL Studio Local AI Foundation

A text-first local VSL builder with:
- plain text raw input and auto slide segmentation
- audio import and microphone recording hooks
- optional faster-whisper transcription
- optional local AI rewrite via Ollama
- optional Piper and XTTS narration backends
- preview + timeline studio pane
- separate editable slide pane opened from timeline interaction
- MP4 export via ffmpeg

Optional Python packages:
    pip install pillow sounddevice soundfile faster-whisper piper-tts TTS

System dependency for video export:
    ffmpeg must be installed and available on PATH, or configured in the app.
"""

from __future__ import annotations

import json
import math
import os
import re
import shutil
import subprocess
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

try:
    from piper import PiperVoice
    PIPER_AVAILABLE = True
except Exception:
    PIPER_AVAILABLE = False
    PiperVoice = None

try:
    from TTS.api import TTS
    XTTS_AVAILABLE = True
except Exception:
    XTTS_AVAILABLE = False
    TTS = None

APP_TITLE = "VSL Studio Local AI Foundation"
PROJECT_EXTENSION = ".vslstudio.json"
DEFAULT_WIDTH = 1280
DEFAULT_HEIGHT = 720
PREVIEW_BG = "#111827"
TIMELINE_BG = "#0f172a"
TIMELINE_GRID = "#1e293b"
TIMELINE_TEXT = "#e5e7eb"
CLIP_FILL = "#2563eb"
CLIP_SELECTED = "#f59e0b"
CLIP_TEXT = "#f8fafc"
TIMELINE_HEADER_WIDTH = 72
TIMELINE_LAYER_HEIGHT = 60
TIMELINE_MIN_DURATION = 0.5
PIXELS_PER_SECOND = 80
BODY_FONT = ("Arial", 12)
TITLE_FONT = ("Arial", 20, "bold")


@dataclass
class TextStyle:
    title_font: str = "Arial"
    body_font: str = "Arial"
    title_size: int = 28
    body_size: int = 18
    title_align: str = "left"
    body_align: str = "left"
    fg_color: str = "#F8FAFC"
    overlay_color: str = "#0F172A"
    overlay_alpha: int = 190


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
    max_chars_per_slide: int = 280
    min_chars_per_slide: int = 80
    tts_backend: str = "piper"
    piper_model: str = ""
    piper_exe: str = ""
    xtts_reference_wav: str = ""
    xtts_language: str = "en"
    use_cuda: bool = False
    whisper_model: str = "small"
    whisper_device: str = "auto"
    ollama_model: str = "qwen3:8b"
    ollama_url: str = "http://localhost:11434/api/generate"
    ffmpeg_path: str = ""
    fps: int = 30
    gap_after_slide_sec: float = 0.35


@dataclass
class Project:
    name: str = "Untitled Project"
    raw_input_text: str = ""
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

    def max_layer(self) -> int:
        return max((s.layer for s in self.slides), default=0)

    def total_duration(self) -> float:
        return max((s.end for s in self.slides), default=0.0)


def project_to_dict(project: Project) -> dict:
    return {
        "name": project.name,
        "raw_input_text": project.raw_input_text,
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
        raw_input_text=data.get("raw_input_text", ""),
        settings=settings,
        slides=slides,
    )
    project.normalize_ids()
    project.sort_slides()
    return project


def plain_project_source(project: Project) -> str:
    lines = [f"#project {project.name}", ""]
    for slide in sorted(project.slides, key=lambda s: (s.start, s.layer, s.slide_id)):
        lines.append(
            f'[[slide id="{slide.slide_id}" title="{slide.title.replace(chr(34), chr(39))}" start="{slide.start:.2f}" duration="{slide.duration:.2f}" layer="{slide.layer}" background="{slide.background.replace(chr(34), chr(39))}" crop_x="{slide.crop_x:.3f}" crop_y="{slide.crop_y:.3f}"]]' 
        )
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


def parse_plain_project_source(text: str, existing_settings: AppSettings, existing_raw_text: str) -> Project:
    project = Project(name="Untitled Project", raw_input_text=existing_raw_text, settings=existing_settings, slides=[])
    lines = text.splitlines()
    current = None
    section = None
    body_lines: List[str] = []
    narr_lines: List[str] = []

    def flush():
        nonlocal current, section, body_lines, narr_lines
        if current is None:
            return
        current.body = "\n".join(body_lines).strip()
        current.narration = "\n".join(narr_lines).strip()
        project.slides.append(current)
        current = None
        section = None
        body_lines = []
        narr_lines = []

    for line in lines:
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
                        narr_lines.append(line)
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
        pieces = re.split(r"(?<=[\.!\?;:])\s+", para)
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
    if len(title) > 52:
        title = title[:49].rstrip() + "..."
    return title.title()


def slides_from_raw_text(raw_text: str, settings: AppSettings, template_style: Optional[TextStyle] = None) -> List[Slide]:
    max_chars = max(80, int(settings.max_chars_per_slide or 280))
    min_chars = max(30, int(settings.min_chars_per_slide or 80))
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
    style_template = template_style or TextStyle()
    for idx, chunk in enumerate(chunks, 1):
        narration = chunk.strip()
        slide = Slide(
            slide_id=f"slide_{idx:03d}",
            title=title_from_chunk(chunk, idx),
            start=round(start, 2),
            duration=estimate_duration_from_text(narration),
            layer=0,
            body=chunk.strip(),
            narration=narration,
            style=TextStyle(**asdict(style_template)),
        )
        slides.append(slide)
        start += slide.duration
    return slides


def estimate_duration_from_text(text: str) -> float:
    words = max(1, len(re.findall(r"\w+", text or "")))
    seconds = max(2.5, words / 2.4)
    return round(seconds, 2)


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
        if path and os.path.isfile(path):
            return path
    return None


def load_font(font_name: str, size: int):
    if not PIL_AVAILABLE:
        return None
    font_file = find_font_file(font_name)
    if font_file:
        try:
            return ImageFont.truetype(font_file, size=size)
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
    odraw.rounded_rectangle((40, height - 265, width - 40, height - 40), radius=18, fill=(r, g, b, max(0, min(255, slide.style.overlay_alpha))))
    odraw.rounded_rectangle((40, 40, width - 40, 160), radius=18, fill=(r, g, b, max(0, min(255, slide.style.overlay_alpha))))
    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")

    draw = ImageDraw.Draw(img)
    title_font = load_font(slide.style.title_font, slide.style.title_size)
    body_font = load_font(slide.style.body_font, slide.style.body_size)
    fill = hex_to_rgb(slide.style.fg_color)
    draw_text_block(draw, (65, 65, width - 130, 80), slide.title, title_font, fill, slide.style.title_align, spacing=6)
    body_text = slide.body.strip() or slide.narration.strip() or "Slide text"
    draw_text_block(draw, (70, height - 240, width - 140, 170), body_text, body_font, fill, slide.style.body_align, spacing=8)
    return img


def wav_duration_seconds(path: Path) -> float:
    with wave.open(str(path), "rb") as wf:
        frames = wf.getnframes()
        rate = wf.getframerate()
        return 0.0 if rate <= 0 else frames / float(rate)


def concatenate_wavs(input_paths: List[Path], output_path: Path, gap_sec: float = 0.0) -> List[Tuple[float, float]]:
    if not input_paths:
        raise ValueError("No input wav files.")
    with wave.open(str(input_paths[0]), "rb") as wf:
        nchannels = wf.getnchannels()
        sampwidth = wf.getsampwidth()
        framerate = wf.getframerate()
        comptype = wf.getcomptype()
        compname = wf.getcompname()
    silence_frames = int(round(gap_sec * framerate))
    silence = b"\x00" * silence_frames * nchannels * sampwidth
    output_path.parent.mkdir(parents=True, exist_ok=True)
    timeline = []
    current = 0.0
    with wave.open(str(output_path), "wb") as out:
        out.setnchannels(nchannels)
        out.setsampwidth(sampwidth)
        out.setframerate(framerate)
        out.setcomptype(comptype, compname)
        for idx, path in enumerate(input_paths):
            with wave.open(str(path), "rb") as wf:
                frames = wf.readframes(wf.getnframes())
                duration = wf.getnframes() / float(framerate)
                start = current
                end = start + duration
                timeline.append((start, end))
                out.writeframes(frames)
                current = end
                if gap_sec > 0 and idx < len(input_paths) - 1:
                    out.writeframes(silence)
                    current += gap_sec
    return timeline


def require_ffmpeg(configured_path: str) -> str:
    if configured_path and os.path.isfile(configured_path):
        return configured_path
    ffmpeg = shutil.which("ffmpeg")
    if ffmpeg:
        return ffmpeg
    raise RuntimeError("ffmpeg was not found. Install ffmpeg or set its path in the Export tab.")


def write_ffmpeg_concat_file(image_paths: List[Path], durations: List[float], output_path: Path) -> None:
    lines = []
    for image_path, duration in zip(image_paths, durations):
        safe = str(image_path.resolve()).replace("'", r"'\\''")
        lines.append(f"file '{safe}'")
        lines.append(f"duration {max(0.05, duration):.4f}")
    safe_last = str(image_paths[-1].resolve()).replace("'", r"'\\''")
    lines.append(f"file '{safe_last}'")
    output_path.write_text("\n".join(lines), encoding="utf-8")


def build_video_from_images_and_audio(image_paths: List[Path], durations: List[float], full_audio_wav: Path, output_mp4: Path, fps: int, ffmpeg_path: str) -> None:
    ffmpeg = require_ffmpeg(ffmpeg_path)
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
        raise RuntimeError(proc.stderr.strip() or "ffmpeg export failed")


class NarrationEngine:
    def __init__(self, settings: AppSettings):
        self.settings = settings
        self._piper_voice = None
        self._xtts = None

    def synthesize(self, text: str, out_wav: Path) -> None:
        backend = (self.settings.tts_backend or "piper").lower()
        if backend == "xtts":
            self._synthesize_xtts(text, out_wav)
        else:
            self._synthesize_piper(text, out_wav)

    def _synthesize_piper(self, text: str, out_wav: Path) -> None:
        if not self.settings.piper_model or not os.path.isfile(self.settings.piper_model):
            raise RuntimeError("Select a valid Piper voice model (.onnx) first.")
        out_wav.parent.mkdir(parents=True, exist_ok=True)
        if PIPER_AVAILABLE:
            if self._piper_voice is None:
                if self.settings.use_cuda:
                    try:
                        self._piper_voice = PiperVoice.load(self.settings.piper_model, use_cuda=True)
                    except Exception:
                        self._piper_voice = PiperVoice.load(self.settings.piper_model)
                else:
                    self._piper_voice = PiperVoice.load(self.settings.piper_model)
            with wave.open(str(out_wav), "wb") as wf:
                self._piper_voice.synthesize_wav(text, wf)
            return
        piper_exe = self.settings.piper_exe.strip()
        if not piper_exe or not os.path.isfile(piper_exe):
            raise RuntimeError("Piper Python package not found and Piper executable path is not configured.")
        cmd = [piper_exe, "--model", self.settings.piper_model, "--output_file", str(out_wav)]
        proc = subprocess.run(cmd, input=text, text=True, capture_output=True)
        if proc.returncode != 0:
            raise RuntimeError(proc.stderr.strip() or "Piper CLI synthesis failed")

    def _synthesize_xtts(self, text: str, out_wav: Path) -> None:
        if not XTTS_AVAILABLE:
            raise RuntimeError("Coqui TTS is not installed. Run: pip install TTS")
        reference = self.settings.xtts_reference_wav.strip()
        if not reference or not os.path.isfile(reference):
            raise RuntimeError("Select a valid XTTS reference voice WAV first.")
        if self._xtts is None:
            device = "cuda" if self.settings.use_cuda else "cpu"
            self._xtts = TTS(model_name="tts_models/multilingual/multi-dataset/xtts_v2", progress_bar=False).to(device)
        self._xtts.tts_to_file(text=text, speaker_wav=reference, language=self.settings.xtts_language or "en", file_path=str(out_wav))


class TranscriptionEngine:
    def __init__(self, settings: AppSettings):
        self.settings = settings
        self._model = None

    def transcribe(self, audio_path: str) -> str:
        if not FASTER_WHISPER_AVAILABLE:
            raise RuntimeError("faster-whisper is not installed. Run: pip install faster-whisper")
        if not os.path.isfile(audio_path):
            raise RuntimeError("Audio file not found.")
        if self._model is None:
            device = self.settings.whisper_device or "auto"
            compute_type = "float16" if self.settings.use_cuda else "int8"
            self._model = WhisperModel(self.settings.whisper_model or "small", device=device, compute_type=compute_type)
        segments, _info = self._model.transcribe(audio_path)
        return " ".join(seg.text.strip() for seg in segments if seg.text.strip()).strip()


class TimelineCanvas(tk.Canvas):
    def __init__(self, master, app: "VSLStudioApp", **kwargs):
        super().__init__(master, bg=TIMELINE_BG, highlightthickness=0, **kwargs)
        self.app = app
        self.drag_mode: Optional[str] = None
        self.drag_slide_id: Optional[str] = None
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
        max_layer = max(4, project.max_layer() + 2)
        total_dur = max(20, math.ceil(project.total_duration()) + 5)
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
            self.create_text(x + 2, 10, anchor="nw", fill="#94a3b8", text=f"{sec}s", font=("Arial", 9))

        for slide in project.slides:
            self.draw_clip(slide)

    def draw_clip(self, slide: Slide) -> None:
        x1 = self.seconds_to_x(slide.start)
        x2 = self.seconds_to_x(slide.end)
        y1 = self.layer_to_y(slide.layer) + 8
        y2 = y1 + TIMELINE_LAYER_HEIGHT - 16
        fill = CLIP_SELECTED if slide.slide_id == self.app.selected_slide_id else CLIP_FILL
        self.create_rectangle(x1, y1, x2, y2, fill=fill, outline="#d1d5db", width=1, tags=("clip", f"slide:{slide.slide_id}"))
        self.create_rectangle(x1, y1, x1 + 7, y2, fill="#1d4ed8", outline="", tags=("left_handle", f"slide:{slide.slide_id}"))
        self.create_rectangle(x2 - 7, y1, x2, y2, fill="#1d4ed8", outline="", tags=("right_handle", f"slide:{slide.slide_id}"))
        label = f"{slide.title} [{slide.duration:.1f}s]"
        self.create_text(x1 + 10, y1 + 8, anchor="nw", fill=CLIP_TEXT, text=label, width=max(50, x2 - x1 - 18), font=("Arial", 10), tags=("clip_text", f"slide:{slide.slide_id}"))

    def find_hit(self, event) -> Tuple[Optional[str], Optional[str]]:
        item = self.find_closest(event.x, event.y)
        if not item:
            return None, None
        tags = self.gettags(item[0])
        slide_id = None
        role = None
        for tag in tags:
            if tag.startswith("slide:"):
                slide_id = tag.split(":", 1)[1]
            elif tag in {"clip", "left_handle", "right_handle", "clip_text"}:
                role = tag
        return slide_id, role

    def on_press(self, event) -> None:
        slide_id, role = self.find_hit(event)
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
        self.app.refresh_editor_from_selected()
        self.app.sync_project_source()
        self.app.timeline.redraw()
        self.app.render_preview()

    def on_release(self, event) -> None:
        self.drag_mode = None
        self.drag_slide_id = None
        self.app.mark_dirty()

    def on_double_click(self, event) -> None:
        slide_id, _role = self.find_hit(event)
        if slide_id:
            self.app.select_slide(slide_id)
            self.app.editor_tabs.select(self.app.slide_editor_tab)
            self.app.title_entry.focus_set()


class VSLStudioApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.current_path: Optional[Path] = None
        self.is_dirty = False
        self.preview_image_cache = None
        self._updating_ui = False
        self._drag_crop = False
        self._recording_thread = None
        self.selected_slide_id: Optional[str] = None
        self.project = self.build_default_project()

        self.build_ui()
        self.select_slide(self.project.slides[0].slide_id if self.project.slides else None)
        self.refresh_raw_input_text()
        self.sync_project_source()
        self.timeline.redraw()
        self.render_preview()
        self.update_title()

    def build_default_project(self) -> Project:
        settings = AppSettings()
        raw = "Welcome to your new text-first VSL workflow. Paste raw words, import audio, or record your voice. Then turn the content into slides, adjust the timeline, and export a narrated video."
        slides = slides_from_raw_text(raw, settings)
        return Project(name="Untitled Project", raw_input_text=raw, settings=settings, slides=slides)

    def build_ui(self) -> None:
        self.root.geometry("1800x1020")
        self.root.minsize(1450, 860)
        self.build_menu()

        outer = ttk.Panedwindow(self.root, orient=tk.HORIZONTAL)
        outer.pack(fill="both", expand=True)

        studio_pane = ttk.Frame(outer)
        editor_pane = ttk.Frame(outer)
        outer.add(studio_pane, weight=3)
        outer.add(editor_pane, weight=2)

        self.build_studio_pane(studio_pane)
        self.build_editor_pane(editor_pane)

        status = ttk.Frame(self.root)
        status.pack(fill="x", side="bottom")
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(status, textvariable=self.status_var).pack(anchor="w", padx=8, pady=4)

    def build_menu(self) -> None:
        menu = tk.Menu(self.root)
        file_menu = tk.Menu(menu, tearoff=False)
        file_menu.add_command(label="New Project", command=self.new_project)
        file_menu.add_command(label="Open Project", command=self.open_project)
        file_menu.add_command(label="Save", command=self.save_project)
        file_menu.add_command(label="Save As", command=self.save_project_as)
        file_menu.add_separator()
        file_menu.add_command(label="Export MP4", command=self.export_mp4_dialog)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menu.add_cascade(label="File", menu=file_menu)

        input_menu = tk.Menu(menu, tearoff=False)
        input_menu.add_command(label="Build Slides From Raw Text", command=self.build_slides_from_raw_text)
        input_menu.add_command(label="Import Audio And Transcribe", command=self.import_audio_and_transcribe)
        input_menu.add_command(label="Record Audio And Transcribe", command=self.record_audio_and_transcribe)
        input_menu.add_command(label="Run Local AI Rewrite", command=self.ai_rewrite_raw_input)
        menu.add_cascade(label="Input", menu=input_menu)

        slide_menu = tk.Menu(menu, tearoff=False)
        slide_menu.add_command(label="Add Slide", command=self.add_slide)
        slide_menu.add_command(label="Duplicate Slide", command=self.duplicate_slide)
        slide_menu.add_command(label="Delete Slide", command=self.delete_slide)
        slide_menu.add_separator()
        slide_menu.add_command(label="Import Background Image", command=self.import_background_image)
        slide_menu.add_command(label="Auto Place Sequentially", command=self.auto_place_sequentially)
        menu.add_cascade(label="Slide", menu=slide_menu)

        self.root.config(menu=menu)

    def build_studio_pane(self, parent) -> None:
        preview_group = ttk.LabelFrame(parent, text="Studio View")
        preview_group.pack(fill="both", expand=True, padx=8, pady=8)
        self.preview_canvas = tk.Canvas(preview_group, bg=PREVIEW_BG, highlightthickness=1, highlightbackground="#334155")
        self.preview_canvas.pack(fill="both", expand=True, padx=8, pady=(8, 6))
        self.preview_canvas.bind("<ButtonPress-1>", self.on_preview_crop_press)
        self.preview_canvas.bind("<B1-Motion>", self.on_preview_crop_drag)
        self.preview_canvas.bind("<ButtonRelease-1>", self.on_preview_crop_release)

        timeline_frame = ttk.LabelFrame(parent, text="Timeline")
        timeline_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.timeline = TimelineCanvas(timeline_frame, self, height=320)
        x_scroll = ttk.Scrollbar(timeline_frame, orient="horizontal", command=self.timeline.xview)
        y_scroll = ttk.Scrollbar(timeline_frame, orient="vertical", command=self.timeline.yview)
        self.timeline.configure(xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)
        self.timeline.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        timeline_frame.rowconfigure(0, weight=1)
        timeline_frame.columnconfigure(0, weight=1)

    def build_editor_pane(self, parent) -> None:
        self.editor_tabs = ttk.Notebook(parent)
        self.editor_tabs.pack(fill="both", expand=True, padx=8, pady=8)

        self.slide_editor_tab = ttk.Frame(self.editor_tabs)
        raw_tab = ttk.Frame(self.editor_tabs)
        source_tab = ttk.Frame(self.editor_tabs)
        settings_tab = ttk.Frame(self.editor_tabs)

        self.editor_tabs.add(self.slide_editor_tab, text="Editable Slide")
        self.editor_tabs.add(raw_tab, text="Raw Input")
        self.editor_tabs.add(source_tab, text="Project Source")
        self.editor_tabs.add(settings_tab, text="AI / Voice / Export")

        self.build_slide_editor_tab(self.slide_editor_tab)
        self.build_raw_input_tab(raw_tab)
        self.build_project_source_tab(source_tab)
        self.build_settings_tab(settings_tab)

    def build_slide_editor_tab(self, parent) -> None:
        top = ttk.Frame(parent)
        top.pack(fill="both", expand=True)

        left = ttk.Frame(top)
        right = ttk.Frame(top)
        left.pack(side="left", fill="both", expand=True, padx=(0, 6))
        right.pack(side="left", fill="y")

        ttk.Label(left, text="Slides").pack(anchor="w", padx=8, pady=(8, 4))
        self.slide_list = tk.Listbox(left, exportselection=False)
        self.slide_list.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.slide_list.bind("<<ListboxSelect>>", self.on_slide_list_select)

        btns = ttk.Frame(left)
        btns.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(btns, text="Add", command=self.add_slide).pack(side="left")
        ttk.Button(btns, text="Duplicate", command=self.duplicate_slide).pack(side="left", padx=4)
        ttk.Button(btns, text="Delete", command=self.delete_slide).pack(side="left", padx=4)

        form = ttk.LabelFrame(right, text="Slide Fields")
        form.pack(fill="x", padx=4, pady=8)

        self.title_var = tk.StringVar()
        self.start_var = tk.StringVar()
        self.duration_var = tk.StringVar()
        self.layer_var = tk.StringVar()
        self.background_var = tk.StringVar()
        self.crop_x_var = tk.StringVar()
        self.crop_y_var = tk.StringVar()

        self.title_font_var = tk.StringVar()
        self.body_font_var = tk.StringVar()
        self.title_size_var = tk.StringVar()
        self.body_size_var = tk.StringVar()
        self.title_align_var = tk.StringVar()
        self.body_align_var = tk.StringVar()
        self.fg_color_var = tk.StringVar()
        self.overlay_color_var = tk.StringVar()
        self.overlay_alpha_var = tk.StringVar()

        row = 0
        self._add_labeled_entry(form, "Title", self.title_var, row); row += 1
        self.title_entry = form.grid_slaves(row=0, column=1)[0]
        self._add_labeled_entry(form, "Start (s)", self.start_var, row); row += 1
        self._add_labeled_entry(form, "Duration (s)", self.duration_var, row); row += 1
        self._add_labeled_entry(form, "Layer", self.layer_var, row); row += 1
        self._add_labeled_entry(form, "Background", self.background_var, row); row += 1
        ttk.Button(form, text="Browse Background", command=self.import_background_image).grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=4); row += 1
        self._add_labeled_entry(form, "Crop X", self.crop_x_var, row); row += 1
        self._add_labeled_entry(form, "Crop Y", self.crop_y_var, row); row += 1
        ttk.Separator(form, orient="horizontal").grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=8); row += 1
        self._add_labeled_entry(form, "Title font", self.title_font_var, row); row += 1
        self._add_labeled_entry(form, "Body font", self.body_font_var, row); row += 1
        self._add_labeled_entry(form, "Title size", self.title_size_var, row); row += 1
        self._add_labeled_entry(form, "Body size", self.body_size_var, row); row += 1
        self._add_labeled_entry(form, "Title align", self.title_align_var, row); row += 1
        self._add_labeled_entry(form, "Body align", self.body_align_var, row); row += 1
        self._add_labeled_entry(form, "Text color", self.fg_color_var, row); row += 1
        self._add_labeled_entry(form, "Overlay color", self.overlay_color_var, row); row += 1
        self._add_labeled_entry(form, "Overlay alpha", self.overlay_alpha_var, row); row += 1

        actions = ttk.Frame(form)
        actions.grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=8)
        ttk.Button(actions, text="Apply To Slide", command=self.apply_editor_to_slide).pack(side="left")
        ttk.Button(actions, text="Apply Style To All", command=self.apply_style_to_all_slides).pack(side="left", padx=6)
        form.columnconfigure(1, weight=1)

        ttk.Label(parent, text="Body").pack(anchor="w", padx=8)
        self.body_text = tk.Text(parent, height=10, wrap="word", font=BODY_FONT)
        self.body_text.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        ttk.Label(parent, text="Narration").pack(anchor="w", padx=8)
        self.narration_text = tk.Text(parent, height=10, wrap="word", font=BODY_FONT)
        self.narration_text.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.body_text.bind("<KeyRelease>", self.on_text_editor_change)
        self.narration_text.bind("<KeyRelease>", self.on_text_editor_change)

    def build_raw_input_tab(self, parent) -> None:
        toolbar = ttk.Frame(parent)
        toolbar.pack(fill="x", padx=8, pady=8)
        ttk.Button(toolbar, text="Build Slides From Raw Text", command=self.build_slides_from_raw_text).pack(side="left")
        ttk.Button(toolbar, text="Import Audio + Transcribe", command=self.import_audio_and_transcribe).pack(side="left", padx=4)
        ttk.Button(toolbar, text="Record + Transcribe", command=self.record_audio_and_transcribe).pack(side="left", padx=4)
        ttk.Button(toolbar, text="Local AI Rewrite", command=self.ai_rewrite_raw_input).pack(side="left", padx=4)
        ttk.Button(toolbar, text="Replace Raw From Slides", command=self.replace_raw_text_from_slides).pack(side="left", padx=4)
        self.raw_input_text = tk.Text(parent, wrap="word", font=("Consolas", 11))
        self.raw_input_text.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def build_project_source_tab(self, parent) -> None:
        toolbar = ttk.Frame(parent)
        toolbar.pack(fill="x", padx=8, pady=8)
        ttk.Button(toolbar, text="Refresh Source", command=self.sync_project_source).pack(side="left")
        ttk.Button(toolbar, text="Rebuild From Source", command=self.rebuild_from_project_source).pack(side="left", padx=4)
        self.project_source = tk.Text(parent, wrap="none", font=("Consolas", 11))
        self.project_source.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def build_settings_tab(self, parent) -> None:
        frm = ttk.Frame(parent)
        frm.pack(fill="both", expand=True, padx=8, pady=8)

        self.canvas_width_var = tk.StringVar()
        self.canvas_height_var = tk.StringVar()
        self.max_chars_var = tk.StringVar()
        self.min_chars_var = tk.StringVar()
        self.tts_backend_var = tk.StringVar()
        self.piper_model_var = tk.StringVar()
        self.piper_exe_var = tk.StringVar()
        self.xtts_reference_var = tk.StringVar()
        self.xtts_language_var = tk.StringVar()
        self.use_cuda_var = tk.BooleanVar(value=False)
        self.whisper_model_var = tk.StringVar()
        self.whisper_device_var = tk.StringVar()
        self.ollama_model_var = tk.StringVar()
        self.ollama_url_var = tk.StringVar()
        self.ffmpeg_path_var = tk.StringVar()
        self.fps_var = tk.StringVar()
        self.gap_var = tk.StringVar()

        row = 0
        self._add_labeled_entry(frm, "Canvas width", self.canvas_width_var, row); row += 1
        self._add_labeled_entry(frm, "Canvas height", self.canvas_height_var, row); row += 1
        self._add_labeled_entry(frm, "Max chars/slide", self.max_chars_var, row); row += 1
        self._add_labeled_entry(frm, "Min chars/slide", self.min_chars_var, row); row += 1
        self._add_labeled_entry(frm, "TTS backend (piper/xtts)", self.tts_backend_var, row); row += 1
        self._add_labeled_entry(frm, "Piper model .onnx", self.piper_model_var, row); row += 1
        ttk.Button(frm, text="Browse Piper Model", command=self.select_piper_model).grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=4); row += 1
        self._add_labeled_entry(frm, "Piper executable", self.piper_exe_var, row); row += 1
        ttk.Button(frm, text="Browse Piper EXE", command=self.select_piper_exe).grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=4); row += 1
        self._add_labeled_entry(frm, "XTTS reference WAV", self.xtts_reference_var, row); row += 1
        ttk.Button(frm, text="Browse XTTS Voice", command=self.select_xtts_reference).grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=4); row += 1
        self._add_labeled_entry(frm, "XTTS language", self.xtts_language_var, row); row += 1
        ttk.Checkbutton(frm, text="Use CUDA if available", variable=self.use_cuda_var).grid(row=row, column=0, columnspan=2, sticky="w", padx=6, pady=4); row += 1
        self._add_labeled_entry(frm, "Whisper model", self.whisper_model_var, row); row += 1
        self._add_labeled_entry(frm, "Whisper device", self.whisper_device_var, row); row += 1
        self._add_labeled_entry(frm, "Ollama model", self.ollama_model_var, row); row += 1
        self._add_labeled_entry(frm, "Ollama URL", self.ollama_url_var, row); row += 1
        self._add_labeled_entry(frm, "ffmpeg path", self.ffmpeg_path_var, row); row += 1
        ttk.Button(frm, text="Browse ffmpeg", command=self.select_ffmpeg).grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=4); row += 1
        self._add_labeled_entry(frm, "Export FPS", self.fps_var, row); row += 1
        self._add_labeled_entry(frm, "Gap after slide (sec)", self.gap_var, row); row += 1

        btns = ttk.Frame(frm)
        btns.grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=8)
        ttk.Button(btns, text="Save Settings To Project", command=self.apply_settings_from_ui).pack(side="left")
        ttk.Button(btns, text="Generate Narration WAVs", command=self.generate_narration_wavs_dialog).pack(side="left", padx=4)
        ttk.Button(btns, text="Export MP4", command=self.export_mp4_dialog).pack(side="left", padx=4)
        frm.columnconfigure(1, weight=1)
        self.refresh_settings_ui()

    def _add_labeled_entry(self, parent, label: str, variable, row: int, width: int = 28) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(parent, textvariable=variable, width=width).grid(row=row, column=1, sticky="ew", padx=6, pady=4)

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

    def set_status(self, text: str) -> None:
        self.status_var.set(text)
        self.root.update_idletasks()

    def get_slide(self, slide_id: Optional[str]) -> Optional[Slide]:
        if not slide_id:
            return None
        for slide in self.project.slides:
            if slide.slide_id == slide_id:
                return slide
        return None

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
        self.refresh_editor_from_selected()
        self.timeline.redraw()
        self.render_preview()

    def refresh_editor_from_selected(self) -> None:
        slide = self.get_slide(self.selected_slide_id)
        self._updating_ui = True
        try:
            if slide is None:
                for var in [self.title_var, self.start_var, self.duration_var, self.layer_var, self.background_var, self.crop_x_var, self.crop_y_var,
                            self.title_font_var, self.body_font_var, self.title_size_var, self.body_size_var, self.title_align_var, self.body_align_var,
                            self.fg_color_var, self.overlay_color_var, self.overlay_alpha_var]:
                    var.set("")
                self.body_text.delete("1.0", tk.END)
                self.narration_text.delete("1.0", tk.END)
                return
            self.title_var.set(slide.title)
            self.start_var.set(f"{slide.start:.2f}")
            self.duration_var.set(f"{slide.duration:.2f}")
            self.layer_var.set(str(slide.layer))
            self.background_var.set(slide.background)
            self.crop_x_var.set(f"{slide.crop_x:.3f}")
            self.crop_y_var.set(f"{slide.crop_y:.3f}")
            self.title_font_var.set(slide.style.title_font)
            self.body_font_var.set(slide.style.body_font)
            self.title_size_var.set(str(slide.style.title_size))
            self.body_size_var.set(str(slide.style.body_size))
            self.title_align_var.set(slide.style.title_align)
            self.body_align_var.set(slide.style.body_align)
            self.fg_color_var.set(slide.style.fg_color)
            self.overlay_color_var.set(slide.style.overlay_color)
            self.overlay_alpha_var.set(str(slide.style.overlay_alpha))
            self.body_text.delete("1.0", tk.END)
            self.body_text.insert("1.0", slide.body)
            self.narration_text.delete("1.0", tk.END)
            self.narration_text.insert("1.0", slide.narration)
        finally:
            self._updating_ui = False

    def refresh_raw_input_text(self) -> None:
        self.raw_input_text.delete("1.0", tk.END)
        self.raw_input_text.insert("1.0", self.project.raw_input_text)

    def refresh_settings_ui(self) -> None:
        s = self.project.settings
        self.canvas_width_var.set(str(s.canvas_width))
        self.canvas_height_var.set(str(s.canvas_height))
        self.max_chars_var.set(str(s.max_chars_per_slide))
        self.min_chars_var.set(str(s.min_chars_per_slide))
        self.tts_backend_var.set(s.tts_backend)
        self.piper_model_var.set(s.piper_model)
        self.piper_exe_var.set(s.piper_exe)
        self.xtts_reference_var.set(s.xtts_reference_wav)
        self.xtts_language_var.set(s.xtts_language)
        self.use_cuda_var.set(bool(s.use_cuda))
        self.whisper_model_var.set(s.whisper_model)
        self.whisper_device_var.set(s.whisper_device)
        self.ollama_model_var.set(s.ollama_model)
        self.ollama_url_var.set(s.ollama_url)
        self.ffmpeg_path_var.set(s.ffmpeg_path)
        self.fps_var.set(str(s.fps))
        self.gap_var.set(str(s.gap_after_slide_sec))

    def apply_settings_from_ui(self) -> None:
        s = self.project.settings
        try:
            s.canvas_width = max(320, int(self.canvas_width_var.get().strip() or DEFAULT_WIDTH))
            s.canvas_height = max(240, int(self.canvas_height_var.get().strip() or DEFAULT_HEIGHT))
            s.max_chars_per_slide = max(80, int(self.max_chars_var.get().strip() or 280))
            s.min_chars_per_slide = max(20, int(self.min_chars_var.get().strip() or 80))
            s.tts_backend = (self.tts_backend_var.get().strip() or "piper").lower()
            s.piper_model = self.piper_model_var.get().strip()
            s.piper_exe = self.piper_exe_var.get().strip()
            s.xtts_reference_wav = self.xtts_reference_var.get().strip()
            s.xtts_language = self.xtts_language_var.get().strip() or "en"
            s.use_cuda = bool(self.use_cuda_var.get())
            s.whisper_model = self.whisper_model_var.get().strip() or "small"
            s.whisper_device = self.whisper_device_var.get().strip() or "auto"
            s.ollama_model = self.ollama_model_var.get().strip() or "qwen3:8b"
            s.ollama_url = self.ollama_url_var.get().strip() or "http://localhost:11434/api/generate"
            s.ffmpeg_path = self.ffmpeg_path_var.get().strip()
            s.fps = max(1, int(self.fps_var.get().strip() or 30))
            s.gap_after_slide_sec = max(0.0, float(self.gap_var.get().strip() or 0.35))
        except ValueError:
            messagebox.showerror("Invalid settings", "One or more setting values are invalid.")
            return
        self.mark_dirty()
        self.set_status("Project settings updated")
        self.render_preview()

    def apply_editor_to_slide(self) -> None:
        slide = self.get_slide(self.selected_slide_id)
        if slide is None:
            return
        try:
            slide.title = self.title_var.get().strip() or "Untitled Slide"
            slide.start = max(0.0, round(float(self.start_var.get().strip() or 0), 2))
            slide.duration = max(TIMELINE_MIN_DURATION, round(float(self.duration_var.get().strip() or 5), 2))
            slide.layer = max(0, int(self.layer_var.get().strip() or 0))
            slide.background = self.background_var.get().strip()
            slide.crop_x = min(1.0, max(0.0, float(self.crop_x_var.get().strip() or 0.5)))
            slide.crop_y = min(1.0, max(0.0, float(self.crop_y_var.get().strip() or 0.5)))
            slide.body = self.body_text.get("1.0", tk.END).strip()
            slide.narration = self.narration_text.get("1.0", tk.END).strip()
            slide.style.title_font = self.title_font_var.get().strip() or "Arial"
            slide.style.body_font = self.body_font_var.get().strip() or "Arial"
            slide.style.title_size = max(10, int(self.title_size_var.get().strip() or 28))
            slide.style.body_size = max(8, int(self.body_size_var.get().strip() or 18))
            slide.style.title_align = (self.title_align_var.get().strip() or "left").lower()
            slide.style.body_align = (self.body_align_var.get().strip() or "left").lower()
            slide.style.fg_color = self.fg_color_var.get().strip() or "#F8FAFC"
            slide.style.overlay_color = self.overlay_color_var.get().strip() or "#0F172A"
            slide.style.overlay_alpha = max(0, min(255, int(self.overlay_alpha_var.get().strip() or 190)))
        except ValueError:
            messagebox.showerror("Invalid value", "One or more slide values are invalid.")
            return
        self.project.sort_slides()
        self.refresh_slide_list()
        self.sync_project_source()
        self.timeline.redraw()
        self.render_preview()
        self.mark_dirty()
        self.set_status("Slide updated")

    def apply_style_to_all_slides(self) -> None:
        slide = self.get_slide(self.selected_slide_id)
        if slide is None:
            return
        self.apply_editor_to_slide()
        template = TextStyle(**asdict(slide.style))
        for target in self.project.slides:
            target.style = TextStyle(**asdict(template))
        self.sync_project_source()
        self.render_preview()
        self.mark_dirty()
        self.set_status("Applied current slide style to all slides")

    def on_slide_list_select(self, event=None) -> None:
        selection = self.slide_list.curselection()
        if not selection:
            return
        idx = selection[0]
        if 0 <= idx < len(self.project.slides):
            self.select_slide(self.project.slides[idx].slide_id)

    def on_text_editor_change(self, event=None) -> None:
        if self._updating_ui:
            return
        slide = self.get_slide(self.selected_slide_id)
        if slide is None:
            return
        slide.body = self.body_text.get("1.0", tk.END).strip()
        slide.narration = self.narration_text.get("1.0", tk.END).strip()
        self.sync_project_source()
        self.render_preview()
        self.mark_dirty()

    def add_slide(self) -> None:
        self.project.normalize_ids()
        idx = len(self.project.slides) + 1
        template_style = self.project.slides[-1].style if self.project.slides else TextStyle()
        start = round(self.project.total_duration(), 2)
        slide = Slide(
            slide_id=f"slide_{idx:03d}",
            title=f"Slide {idx}",
            start=start,
            duration=5.0,
            layer=0,
            body="Add body text here.",
            narration="Add narration here.",
            style=TextStyle(**asdict(template_style)),
        )
        self.project.slides.append(slide)
        self.project.sort_slides()
        self.select_slide(slide.slide_id)
        self.sync_project_source()
        self.mark_dirty()
        self.set_status("Slide added")

    def duplicate_slide(self) -> None:
        slide = self.get_slide(self.selected_slide_id)
        if slide is None:
            return
        new_slide = slide.clone()
        new_slide.slide_id = ""
        new_slide.title = f"{slide.title} Copy"
        new_slide.start = round(slide.end + 0.5, 2)
        self.project.slides.append(new_slide)
        self.project.normalize_ids()
        self.project.sort_slides()
        self.select_slide(new_slide.slide_id)
        self.sync_project_source()
        self.mark_dirty()
        self.set_status("Slide duplicated")

    def delete_slide(self) -> None:
        slide = self.get_slide(self.selected_slide_id)
        if slide is None:
            return
        if not messagebox.askyesno("Delete slide", f"Delete '{slide.title}'?"):
            return
        self.project.slides = [s for s in self.project.slides if s.slide_id != slide.slide_id]
        self.project.normalize_ids()
        next_sel = self.project.slides[0].slide_id if self.project.slides else None
        self.select_slide(next_sel)
        self.sync_project_source()
        self.mark_dirty()
        self.set_status("Slide deleted")

    def auto_place_sequentially(self) -> None:
        t = 0.0
        for slide in sorted(self.project.slides, key=lambda s: (s.start, s.layer, s.slide_id)):
            slide.start = round(t, 2)
            slide.layer = 0
            t += slide.duration
        self.project.sort_slides()
        self.refresh_slide_list()
        self.timeline.redraw()
        self.render_preview()
        self.sync_project_source()
        self.mark_dirty()
        self.set_status("Slides auto-placed sequentially")

    def import_background_image(self) -> None:
        slide = self.get_slide(self.selected_slide_id)
        if slide is None:
            return
        path = filedialog.askopenfilename(title="Select background image", filetypes=[("Images", "*.png *.jpg *.jpeg *.webp *.bmp"), ("All files", "*.*")])
        if not path:
            return
        slide.background = path
        self.background_var.set(path)
        self.sync_project_source()
        self.render_preview()
        self.mark_dirty()
        self.set_status("Background image linked to slide")

    def replace_raw_text_from_slides(self) -> None:
        text = "\n\n".join((s.narration.strip() or s.body.strip() or s.title.strip()) for s in self.project.slides if (s.narration.strip() or s.body.strip() or s.title.strip()))
        self.project.raw_input_text = text
        self.refresh_raw_input_text()
        self.mark_dirty()
        self.set_status("Raw input replaced from current slides")

    def build_slides_from_raw_text(self) -> None:
        self.apply_settings_from_ui()
        raw = self.raw_input_text.get("1.0", tk.END).strip()
        if not raw:
            messagebox.showinfo("No raw text", "Paste or transcribe text first.")
            return
        template = self.get_slide(self.selected_slide_id).style if self.get_slide(self.selected_slide_id) else TextStyle()
        slides = slides_from_raw_text(raw, self.project.settings, template)
        if not slides:
            messagebox.showerror("Build failed", "Could not generate slides from the raw text.")
            return
        self.project.raw_input_text = raw
        self.project.slides = slides
        self.project.normalize_ids()
        self.project.sort_slides()
        self.select_slide(self.project.slides[0].slide_id if self.project.slides else None)
        self.sync_project_source()
        self.mark_dirty()
        self.set_status(f"Built {len(self.project.slides)} slides from raw text")

    def import_audio_and_transcribe(self) -> None:
        path = filedialog.askopenfilename(title="Select audio or video file", filetypes=[("Media", "*.wav *.mp3 *.m4a *.flac *.mp4 *.mov *.mkv"), ("All files", "*.*")])
        if not path:
            return
        self.apply_settings_from_ui()
        self.run_in_thread(lambda: self._transcribe_file_worker(path), "Transcribing imported media...")

    def _transcribe_file_worker(self, path: str) -> None:
        engine = TranscriptionEngine(self.project.settings)
        transcript = engine.transcribe(path)
        self.root.after(0, lambda: self._append_transcript_to_raw_input(transcript, source_label=os.path.basename(path)))

    def record_audio_and_transcribe(self) -> None:
        if not (SOUNDDEVICE_AVAILABLE and SOUNDFILE_AVAILABLE):
            messagebox.showerror("Recording unavailable", "Install sounddevice and soundfile first.\n\npip install sounddevice soundfile")
            return
        seconds = simpledialog.askinteger("Record audio", "Record how many seconds?", initialvalue=20, minvalue=1, maxvalue=600)
        if not seconds:
            return
        out_dir = Path.home() / "VSLStudio_Recordings"
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"recording_{int(time.time())}.wav"
        self.apply_settings_from_ui()
        self.run_in_thread(lambda: self._record_and_transcribe_worker(seconds, out_path), "Recording and transcribing audio...")

    def _record_and_transcribe_worker(self, seconds: int, out_path: Path) -> None:
        fs = 16000
        audio = sd.rec(int(seconds * fs), samplerate=fs, channels=1, dtype="float32")
        sd.wait()
        sf.write(str(out_path), audio, fs)
        engine = TranscriptionEngine(self.project.settings)
        transcript = engine.transcribe(str(out_path))
        self.root.after(0, lambda: self._append_transcript_to_raw_input(transcript, source_label=out_path.name))

    def _append_transcript_to_raw_input(self, transcript: str, source_label: str = "transcript") -> None:
        if not transcript.strip():
            messagebox.showwarning("No transcript", "Transcription completed but no text was returned.")
            return
        current = self.raw_input_text.get("1.0", tk.END).strip()
        add = transcript.strip()
        combined = f"{current}\n\n{add}".strip() if current else add
        self.project.raw_input_text = combined
        self.refresh_raw_input_text()
        self.editor_tabs.select(1)
        self.mark_dirty()
        self.set_status(f"Added transcript from {source_label}")

    def ai_rewrite_raw_input(self) -> None:
        self.apply_settings_from_ui()
        raw = self.raw_input_text.get("1.0", tk.END).strip()
        if not raw:
            messagebox.showinfo("No input", "Paste or transcribe content first.")
            return
        prompt = (
            "Rewrite the following into concise, slide-friendly VSL content. Preserve the meaning. "
            "Use plain language, natural spoken rhythm, and paragraph breaks suitable for slide segmentation. "
            "Do not use markdown bullets unless truly necessary.\n\n"
            + raw
        )
        self.run_in_thread(lambda: self._ollama_rewrite_worker(prompt), "Running local AI rewrite...")

    def _ollama_rewrite_worker(self, prompt: str) -> None:
        payload = json.dumps({"model": self.project.settings.ollama_model, "prompt": prompt, "stream": False}).encode("utf-8")
        req = urllib.request.Request(self.project.settings.ollama_url, data=payload, headers={"Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=300) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except urllib.error.URLError as exc:
            self.root.after(0, lambda: messagebox.showerror("Ollama error", f"Could not reach the local Ollama server.\n\n{exc}"))
            return
        response_text = (data.get("response") or "").strip()
        self.root.after(0, lambda: self._replace_raw_with_ai_text(response_text))

    def _replace_raw_with_ai_text(self, text: str) -> None:
        if not text:
            messagebox.showwarning("No AI output", "The local AI returned no text.")
            return
        self.project.raw_input_text = text
        self.refresh_raw_input_text()
        self.mark_dirty()
        self.set_status("Raw text updated from local AI")

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

    def sync_project_source(self) -> None:
        source = plain_project_source(self.project)
        self.project_source.delete("1.0", tk.END)
        self.project_source.insert("1.0", source)

    def rebuild_from_project_source(self) -> None:
        try:
            self.project = parse_plain_project_source(self.project_source.get("1.0", tk.END), self.project.settings, self.project.raw_input_text)
            self.selected_slide_id = self.project.slides[0].slide_id if self.project.slides else None
            self.refresh_slide_list()
            self.refresh_editor_from_selected()
            self.timeline.redraw()
            self.render_preview()
            self.mark_dirty()
            self.set_status("Rebuilt slides from plain project source")
        except Exception as exc:
            messagebox.showerror("Parse error", f"Could not rebuild project source.\n\n{exc}")

    def render_preview(self) -> None:
        self.preview_canvas.delete("all")
        slide = self.get_slide(self.selected_slide_id)
        canvas_w = max(1, self.preview_canvas.winfo_width())
        canvas_h = max(1, self.preview_canvas.winfo_height())
        self.preview_canvas.create_rectangle(0, 0, canvas_w, canvas_h, fill=PREVIEW_BG, outline="")
        if slide is None:
            self.preview_canvas.create_text(canvas_w / 2, canvas_h / 2, text="No slide selected", fill="#cbd5e1", font=TITLE_FONT)
            return
        if PIL_AVAILABLE:
            img = render_slide_to_image(slide, (canvas_w, canvas_h))
            if img is not None:
                self.preview_image_cache = ImageTk.PhotoImage(img)
                self.preview_canvas.create_image(0, 0, image=self.preview_image_cache, anchor="nw")
        else:
            self.preview_canvas.create_text(canvas_w / 2, canvas_h / 2, text="Install pillow for slide preview", fill="#cbd5e1", font=TITLE_FONT)
        self.preview_canvas.create_text(canvas_w - 12, 10, anchor="ne", text=f"{slide.start:.2f}s • {slide.duration:.2f}s • L{slide.layer}", fill="#cbd5e1", font=("Arial", 10))

    def on_preview_crop_press(self, event) -> None:
        slide = self.get_slide(self.selected_slide_id)
        if slide is None:
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
        self.crop_x_var.set(f"{slide.crop_x:.3f}")
        self.crop_y_var.set(f"{slide.crop_y:.3f}")
        self.render_preview()
        self.sync_project_source()

    def select_piper_model(self) -> None:
        path = filedialog.askopenfilename(title="Select Piper voice model", filetypes=[("ONNX", "*.onnx"), ("All files", "*.*")])
        if path:
            self.piper_model_var.set(path)

    def select_piper_exe(self) -> None:
        path = filedialog.askopenfilename(title="Select Piper executable", filetypes=[("Executable", "*.exe"), ("All files", "*.*")])
        if path:
            self.piper_exe_var.set(path)

    def select_xtts_reference(self) -> None:
        path = filedialog.askopenfilename(title="Select XTTS reference voice WAV", filetypes=[("WAV", "*.wav"), ("Audio", "*.wav *.mp3 *.m4a"), ("All files", "*.*")])
        if path:
            self.xtts_reference_var.set(path)

    def select_ffmpeg(self) -> None:
        path = filedialog.askopenfilename(title="Select ffmpeg executable", filetypes=[("Executable", "*.exe"), ("All files", "*.*")])
        if path:
            self.ffmpeg_path_var.set(path)

    def export_mp4_dialog(self) -> None:
        if not self.project.slides:
            messagebox.showerror("No slides", "There are no slides to export.")
            return
        self.apply_settings_from_ui()
        path = filedialog.asksaveasfilename(title="Export MP4", defaultextension=".mp4", filetypes=[("MP4", "*.mp4"), ("All files", "*.*")])
        if not path:
            return
        self.run_in_thread(lambda: self._export_mp4_worker(Path(path)), "Exporting narrated MP4...")

    def _export_mp4_worker(self, output_mp4: Path) -> None:
        if not PIL_AVAILABLE:
            raise RuntimeError("Pillow is required for slide rendering. Run: pip install pillow")
        temp_dir = output_mp4.parent / f"{output_mp4.stem}_export"
        images_dir = temp_dir / "images"
        audio_dir = temp_dir / "audio"
        images_dir.mkdir(parents=True, exist_ok=True)
        audio_dir.mkdir(parents=True, exist_ok=True)
        engine = NarrationEngine(self.project.settings)
        image_paths: List[Path] = []
        wav_paths: List[Path] = []
        durations: List[float] = []

        for idx, slide in enumerate(self.project.slides, 1):
            self.root.after(0, lambda i=idx: self.set_status(f"Rendering slide {i}/{len(self.project.slides)}..."))
            img = render_slide_to_image(slide, (self.project.settings.canvas_width, self.project.settings.canvas_height))
            if img is None:
                raise RuntimeError("Pillow is required for slide rendering.")
            img_path = images_dir / f"slide_{idx:03d}.png"
            img.save(img_path)
            image_paths.append(img_path)

            narration_text = slide.narration.strip() or slide.body.strip() or slide.title.strip() or f"Slide {idx}"
            wav_path = audio_dir / f"slide_{idx:03d}.wav"
            self.root.after(0, lambda i=idx: self.set_status(f"Generating narration {i}/{len(self.project.slides)}..."))
            engine.synthesize(narration_text, wav_path)
            wav_paths.append(wav_path)
            durations.append(max(slide.duration, wav_duration_seconds(wav_path) + self.project.settings.gap_after_slide_sec))

        full_audio = temp_dir / "full_audio.wav"
        concatenate_wavs(wav_paths, full_audio, gap_sec=self.project.settings.gap_after_slide_sec)
        build_video_from_images_and_audio(image_paths, durations, full_audio, output_mp4, self.project.settings.fps, self.project.settings.ffmpeg_path)
        self.root.after(0, lambda: self.set_status(f"Exported MP4: {output_mp4}"))

    def generate_narration_wavs_dialog(self) -> None:
        if not self.project.slides:
            messagebox.showerror("No slides", "There are no slides to narrate.")
            return
        self.apply_settings_from_ui()
        out_dir = filedialog.askdirectory(title="Select output folder for narration WAVs")
        if not out_dir:
            return
        self.run_in_thread(lambda: self._generate_wavs_worker(Path(out_dir)), "Generating narration WAVs...")

    def _generate_wavs_worker(self, out_dir: Path) -> None:
        out_dir.mkdir(parents=True, exist_ok=True)
        engine = NarrationEngine(self.project.settings)
        for idx, slide in enumerate(self.project.slides, 1):
            text = slide.narration.strip() or slide.body.strip() or slide.title.strip()
            if not text:
                text = f"Slide {idx}"
            self.root.after(0, lambda i=idx: self.set_status(f"Generating narration {i}/{len(self.project.slides)}..."))
            engine.synthesize(text, out_dir / f"slide_{idx:03d}.wav")
        self.root.after(0, lambda: self.set_status(f"Saved narration WAVs to: {out_dir}"))

    def new_project(self) -> None:
        if not self.confirm_discard_changes():
            return
        self.project = self.build_default_project()
        self.current_path = None
        self.selected_slide_id = self.project.slides[0].slide_id if self.project.slides else None
        self.refresh_raw_input_text()
        self.refresh_settings_ui()
        self.refresh_slide_list()
        self.refresh_editor_from_selected()
        self.sync_project_source()
        self.timeline.redraw()
        self.render_preview()
        self.mark_clean()
        self.set_status("New project created")

    def open_project(self) -> None:
        if not self.confirm_discard_changes():
            return
        path = filedialog.askopenfilename(title="Open project", filetypes=[("VSL Studio Project", f"*{PROJECT_EXTENSION}"), ("JSON", "*.json"), ("All files", "*.*")])
        if not path:
            return
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        self.project = project_from_dict(data)
        self.current_path = Path(path)
        self.selected_slide_id = self.project.slides[0].slide_id if self.project.slides else None
        self.refresh_raw_input_text()
        self.refresh_settings_ui()
        self.refresh_slide_list()
        self.refresh_editor_from_selected()
        self.sync_project_source()
        self.timeline.redraw()
        self.render_preview()
        self.mark_clean()
        self.set_status(f"Opened: {path}")

    def save_project(self) -> None:
        if self.current_path is None:
            self.save_project_as()
            return
        self.project.raw_input_text = self.raw_input_text.get("1.0", tk.END).strip()
        self.apply_settings_from_ui()
        self.apply_editor_to_slide()
        self.current_path.write_text(json.dumps(project_to_dict(self.project), indent=2), encoding="utf-8")
        self.mark_clean()
        self.set_status(f"Saved: {self.current_path}")

    def save_project_as(self) -> None:
        path = filedialog.asksaveasfilename(title="Save project", defaultextension=PROJECT_EXTENSION, filetypes=[("VSL Studio Project", f"*{PROJECT_EXTENSION}"), ("JSON", "*.json"), ("All files", "*.*")])
        if not path:
            return
        self.current_path = Path(path)
        self.save_project()

    def confirm_discard_changes(self) -> bool:
        if not self.is_dirty:
            return True
        return messagebox.askyesno("Unsaved changes", "Discard unsaved changes?")


def main() -> None:
    root = tk.Tk()
    try:
        style = ttk.Style(root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass
    VSLStudioApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
