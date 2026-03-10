#!/usr/bin/env python3
"""
Standalone VSL Timeline Foundation

A plain-text driven slide editor and timeline foundation for building VSL-style
presentations without any PowerPoint dependency.

What this foundation includes:
- Plain text project format using easy [[slide]] blocks
- Open / save / new project
- Slide list
- Structured slide editor (title, duration, layer, body, narration, background)
- Two-way sync between text format and slide objects
- Preview canvas with optional background image
- Timeline canvas with draggable clips
- Trim left / trim right by dragging clip edges
- Move clips across time and layers

What this foundation intentionally leaves for later:
- Audio transcription
- TTS / voice cloning
- MP4 rendering
- Multi-track audio waveforms
- Keyframed animation

Dependencies:
- Python 3.10+
- tkinter (usually bundled)
- Pillow optional for image preview: pip install pillow
"""

from __future__ import annotations

import math
import os
import re
import textwrap
import tkinter as tk
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional, Tuple

try:
    from PIL import Image, ImageOps, ImageTk
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False
    Image = None
    ImageOps = None
    ImageTk = None

APP_TITLE = "VSL Timeline Foundation"
DEFAULT_WIDTH = 1280
DEFAULT_HEIGHT = 720
PIXELS_PER_SECOND = 70
TIMELINE_LAYER_HEIGHT = 60
TIMELINE_HEADER_WIDTH = 70
TIMELINE_MIN_DURATION = 0.5
PREVIEW_BG = "#111827"
TIMELINE_BG = "#0f172a"
TIMELINE_GRID = "#1e293b"
TIMELINE_TEXT = "#e5e7eb"
CLIP_FILL = "#2563eb"
CLIP_SELECTED = "#f59e0b"
CLIP_TEXT = "#f8fafc"
BODY_FONT = ("Arial", 12)
TITLE_FONT = ("Arial", 20, "bold")
SECTION_RE = re.compile(r"^\s*(body|narration)\s*:\s*$", re.IGNORECASE)
SLIDE_OPEN_RE = re.compile(r"^\s*\[\[slide\s*(.*?)\]\]\s*$", re.IGNORECASE)
SLIDE_CLOSE_RE = re.compile(r"^\s*\[\[/slide\]\]\s*$", re.IGNORECASE)
ATTR_RE = re.compile(r"(\w+)\s*=\s*\"(.*?)\"")


@dataclass
class Slide:
    slide_id: str
    title: str = "Untitled Slide"
    duration: float = 5.0
    layer: int = 0
    start: float = 0.0
    body: str = ""
    narration: str = ""
    background: str = ""
    crop_x: float = 0.5
    crop_y: float = 0.5
    notes: str = ""

    def clone(self) -> "Slide":
        return Slide(
            slide_id=self.slide_id,
            title=self.title,
            duration=self.duration,
            layer=self.layer,
            start=self.start,
            body=self.body,
            narration=self.narration,
            background=self.background,
            crop_x=self.crop_x,
            crop_y=self.crop_y,
            notes=self.notes,
        )

    @property
    def end(self) -> float:
        return self.start + self.duration


@dataclass
class Project:
    name: str = "Untitled Project"
    width: int = DEFAULT_WIDTH
    height: int = DEFAULT_HEIGHT
    slides: List[Slide] = field(default_factory=list)

    def sort_slides(self) -> None:
        self.slides.sort(key=lambda s: (s.start, s.layer, s.slide_id))

    def normalize_ids(self) -> None:
        seen = set()
        for i, slide in enumerate(self.slides, 1):
            if not slide.slide_id or slide.slide_id in seen:
                slide.slide_id = f"slide_{i:03d}"
            seen.add(slide.slide_id)

    def total_duration(self) -> float:
        if not self.slides:
            return 0.0
        return max(slide.end for slide in self.slides)

    def max_layer(self) -> int:
        if not self.slides:
            return 0
        return max(slide.layer for slide in self.slides)


def parse_attrs(attr_text: str) -> dict:
    attrs = {}
    for key, value in ATTR_RE.findall(attr_text):
        attrs[key.lower()] = value
    return attrs


def parse_project_text(text: str) -> Project:
    lines = text.splitlines()
    project = Project()
    idx = 0

    current_slide: Optional[Slide] = None
    current_section: Optional[str] = None
    body_lines: List[str] = []
    narration_lines: List[str] = []

    def finalize_slide() -> None:
        nonlocal current_slide, current_section, body_lines, narration_lines, project
        if current_slide is None:
            return
        current_slide.body = "\n".join(body_lines).strip()
        current_slide.narration = "\n".join(narration_lines).strip()
        project.slides.append(current_slide)
        current_slide = None
        current_section = None
        body_lines = []
        narration_lines = []

    while idx < len(lines):
        line = lines[idx]
        stripped = line.strip()

        if current_slide is None:
            if stripped.startswith("#project"):
                parts = stripped.split(maxsplit=1)
                project.name = parts[1].strip() if len(parts) > 1 else project.name
            elif stripped.startswith("canvas:"):
                value = stripped.split(":", 1)[1].strip().lower()
                if "x" in value:
                    w, h = value.split("x", 1)
                    try:
                        project.width = int(w)
                        project.height = int(h)
                    except ValueError:
                        pass
            else:
                m = SLIDE_OPEN_RE.match(line)
                if m:
                    attrs = parse_attrs(m.group(1))
                    current_slide = Slide(
                        slide_id=attrs.get("id", f"slide_{len(project.slides)+1:03d}"),
                        title=attrs.get("title", "Untitled Slide"),
                        duration=max(TIMELINE_MIN_DURATION, float(attrs.get("duration", "5") or 5)),
                        layer=max(0, int(attrs.get("layer", "0") or 0)),
                        start=max(0.0, float(attrs.get("start", "0") or 0)),
                        background=attrs.get("background", ""),
                        crop_x=min(1.0, max(0.0, float(attrs.get("crop_x", "0.5") or 0.5))),
                        crop_y=min(1.0, max(0.0, float(attrs.get("crop_y", "0.5") or 0.5))),
                        notes=attrs.get("notes", ""),
                    )
        else:
            if SLIDE_CLOSE_RE.match(line):
                finalize_slide()
            else:
                sec = SECTION_RE.match(line)
                if sec:
                    current_section = sec.group(1).lower()
                else:
                    if current_section == "body":
                        body_lines.append(line)
                    elif current_section == "narration":
                        narration_lines.append(line)
                    else:
                        pass
        idx += 1

    finalize_slide()
    project.sort_slides()
    project.normalize_ids()
    return project


def project_to_text(project: Project) -> str:
    project.sort_slides()
    project.normalize_ids()
    out = [f"#project {project.name}", f"canvas: {project.width}x{project.height}", ""]
    for slide in project.slides:
        attrs = (
            f'id="{slide.slide_id}" '
            f'title="{escape_attr(slide.title)}" '
            f'duration="{slide.duration:.2f}" '
            f'layer="{slide.layer}" '
            f'start="{slide.start:.2f}" '
            f'background="{escape_attr(slide.background)}" '
            f'crop_x="{slide.crop_x:.3f}" '
            f'crop_y="{slide.crop_y:.3f}"'
        )
        out.append(f"[[slide {attrs}]]")
        out.append("body:")
        out.append(slide.body.rstrip())
        out.append("")
        out.append("narration:")
        out.append(slide.narration.rstrip())
        out.append("[[/slide]]")
        out.append("")
    return "\n".join(out).rstrip() + "\n"


def escape_attr(value: str) -> str:
    return (value or "").replace('"', "'")


class TimelineCanvas(tk.Canvas):
    def __init__(self, master, app: "VSLFoundationApp", **kwargs):
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

    def x_to_seconds(self, x: float) -> float:
        return max(0.0, (x - TIMELINE_HEADER_WIDTH) / PIXELS_PER_SECOND)

    def layer_to_y(self, layer: int) -> float:
        return layer * TIMELINE_LAYER_HEIGHT

    def y_to_layer(self, y: float) -> int:
        return max(0, int(y // TIMELINE_LAYER_HEIGHT))

    def redraw(self) -> None:
        self.delete("all")
        project = self.app.project
        max_layer = max(4, project.max_layer() + 2)
        total_dur = max(20, math.ceil(project.total_duration()) + 5)
        width = int(self.seconds_to_x(total_dur)) + 200
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
        self.create_rectangle(x1, y1, x1 + 6, y2, fill="#1d4ed8", outline="", tags=("left_handle", f"slide:{slide.slide_id}"))
        self.create_rectangle(x2 - 6, y1, x2, y2, fill="#1d4ed8", outline="", tags=("right_handle", f"slide:{slide.slide_id}"))
        label = f"{slide.title} [{slide.duration:.1f}s]"
        self.create_text(x1 + 10, y1 + 8, anchor="nw", fill=CLIP_TEXT, text=label, width=max(50, x2 - x1 - 18), font=("Arial", 10), tags=("clip_text", f"slide:{slide.slide_id}"))

    def find_slide_id_at(self, event) -> Tuple[Optional[str], Optional[str]]:
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
        slide_id, role = self.find_slide_id_at(event)
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
            new_duration = max(TIMELINE_MIN_DURATION, round(self.orig_duration + dx_sec, 2))
            slide.duration = new_duration

        self.app.refresh_editor_fields_from_selected(preserve_text_focus=False)
        self.app.timeline.redraw()
        self.app.render_preview()
        self.app.refresh_text_from_project(mark_clean=False)

    def on_release(self, event) -> None:
        self.drag_mode = None
        self.drag_slide_id = None
        self.app.mark_dirty()

    def on_double_click(self, event) -> None:
        slide_id, role = self.find_slide_id_at(event)
        if slide_id:
            self.app.select_slide(slide_id)
            self.app.focus_title_entry()


class VSLFoundationApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.project = Project(
            slides=[
                Slide(slide_id="slide_001", title="Intro", start=0.0, duration=5.0, layer=0, body="Opening headline and visual support.", narration="Welcome to your new plain text driven VSL workflow."),
                Slide(slide_id="slide_002", title="Problem", start=5.5, duration=6.5, layer=0, body="State the problem with clarity.", narration="Most slideshow tools make fast iteration harder than it should be."),
            ]
        )
        self.project.normalize_ids()
        self.current_path: Optional[Path] = None
        self.selected_slide_id: Optional[str] = None
        self.preview_image_cache = None
        self.is_dirty = False
        self._updating_ui = False
        self._drag_crop = False

        self.build_ui()
        self.select_slide(self.project.slides[0].slide_id if self.project.slides else None)
        self.refresh_text_from_project(mark_clean=True)
        self.timeline.redraw()
        self.render_preview()

    def build_ui(self) -> None:
        self.root.geometry("1600x950")
        self.root.minsize(1300, 800)
        self.build_menu()

        outer = ttk.Panedwindow(self.root, orient=tk.HORIZONTAL)
        outer.pack(fill="both", expand=True)

        left = ttk.Frame(outer)
        center = ttk.Frame(outer)
        right = ttk.Frame(outer)
        outer.add(left, weight=1)
        outer.add(center, weight=2)
        outer.add(right, weight=2)

        self.build_left_panel(left)
        self.build_center_panel(center)
        self.build_right_panel(right)

        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill="x", side="bottom")
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(status_frame, textvariable=self.status_var).pack(anchor="w", padx=8, pady=4)

    def build_menu(self) -> None:
        menu = tk.Menu(self.root)
        file_menu = tk.Menu(menu, tearoff=False)
        file_menu.add_command(label="New Project", command=self.new_project)
        file_menu.add_command(label="Open Project", command=self.open_project)
        file_menu.add_command(label="Save", command=self.save_project)
        file_menu.add_command(label="Save As", command=self.save_project_as)
        file_menu.add_separator()
        file_menu.add_command(label="Export Plain Text", command=self.export_plain_text)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menu.add_cascade(label="File", menu=file_menu)

        slide_menu = tk.Menu(menu, tearoff=False)
        slide_menu.add_command(label="Add Slide", command=self.add_slide)
        slide_menu.add_command(label="Duplicate Slide", command=self.duplicate_slide)
        slide_menu.add_command(label="Delete Slide", command=self.delete_slide)
        slide_menu.add_separator()
        slide_menu.add_command(label="Import Background Image", command=self.import_background_image)
        menu.add_cascade(label="Slide", menu=slide_menu)

        text_menu = tk.Menu(menu, tearoff=False)
        text_menu.add_command(label="Rebuild Project From Text", command=self.rebuild_project_from_text)
        text_menu.add_command(label="Refresh Text From Slides", command=lambda: self.refresh_text_from_project(mark_clean=False))
        menu.add_cascade(label="Text", menu=text_menu)

        self.root.config(menu=menu)

    def build_left_panel(self, parent) -> None:
        ttk.Label(parent, text="Slides").pack(anchor="w", padx=8, pady=(8, 4))
        self.slide_list = tk.Listbox(parent, exportselection=False)
        self.slide_list.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.slide_list.bind("<<ListboxSelect>>", self.on_slide_list_select)

        btn_row = ttk.Frame(parent)
        btn_row.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(btn_row, text="Add", command=self.add_slide).pack(side="left", padx=(0, 4))
        ttk.Button(btn_row, text="Duplicate", command=self.duplicate_slide).pack(side="left", padx=4)
        ttk.Button(btn_row, text="Delete", command=self.delete_slide).pack(side="left", padx=4)

    def build_center_panel(self, parent) -> None:
        top = ttk.Frame(parent)
        top.pack(fill="both", expand=True)

        ttk.Label(top, text="Slide Editor").pack(anchor="w", padx=8, pady=(8, 4))
        form = ttk.Frame(top)
        form.pack(fill="x", padx=8)

        self.title_var = tk.StringVar()
        self.duration_var = tk.StringVar()
        self.start_var = tk.StringVar()
        self.layer_var = tk.StringVar()
        self.background_var = tk.StringVar()
        self.crop_x_var = tk.StringVar()
        self.crop_y_var = tk.StringVar()

        row = 0
        ttk.Label(form, text="Title").grid(row=row, column=0, sticky="w", pady=4)
        self.title_entry = ttk.Entry(form, textvariable=self.title_var)
        self.title_entry.grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        ttk.Label(form, text="Start (s)").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=self.start_var).grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        ttk.Label(form, text="Duration (s)").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=self.duration_var).grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        ttk.Label(form, text="Layer").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=self.layer_var).grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        ttk.Label(form, text="Background").grid(row=row, column=0, sticky="w", pady=4)
        bg_row = ttk.Frame(form)
        bg_row.grid(row=row, column=1, sticky="ew", pady=4)
        ttk.Entry(bg_row, textvariable=self.background_var).pack(side="left", fill="x", expand=True)
        ttk.Button(bg_row, text="Browse", command=self.import_background_image).pack(side="left", padx=(4, 0))
        row += 1

        ttk.Label(form, text="Crop X (0-1)").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=self.crop_x_var).grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        ttk.Label(form, text="Crop Y (0-1)").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=self.crop_y_var).grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        form.columnconfigure(1, weight=1)

        editor_buttons = ttk.Frame(top)
        editor_buttons.pack(fill="x", padx=8, pady=(4, 8))
        ttk.Button(editor_buttons, text="Apply Slide Changes", command=self.apply_editor_to_slide).pack(side="left")
        ttk.Button(editor_buttons, text="Auto Place Sequentially", command=self.auto_place_sequentially).pack(side="left", padx=6)

        ttk.Label(top, text="Body").pack(anchor="w", padx=8)
        self.body_text = tk.Text(top, height=10, wrap="word", font=BODY_FONT)
        self.body_text.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        ttk.Label(top, text="Narration").pack(anchor="w", padx=8)
        self.narration_text = tk.Text(top, height=10, wrap="word", font=BODY_FONT)
        self.narration_text.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.body_text.bind("<KeyRelease>", self.on_text_editor_change)
        self.narration_text.bind("<KeyRelease>", self.on_text_editor_change)
        for variable in [self.title_var, self.duration_var, self.start_var, self.layer_var, self.background_var, self.crop_x_var, self.crop_y_var]:
            variable.trace_add("write", self.on_field_change)

    def build_right_panel(self, parent) -> None:
        ttk.Label(parent, text="Preview").pack(anchor="w", padx=8, pady=(8, 4))
        preview_frame = ttk.Frame(parent)
        preview_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.preview_canvas = tk.Canvas(preview_frame, bg=PREVIEW_BG, width=640, height=360, highlightthickness=1, highlightbackground="#334155")
        self.preview_canvas.pack(fill="both", expand=True)
        self.preview_canvas.bind("<ButtonPress-1>", self.on_preview_crop_press)
        self.preview_canvas.bind("<B1-Motion>", self.on_preview_crop_drag)
        self.preview_canvas.bind("<ButtonRelease-1>", self.on_preview_crop_release)

        ttk.Label(parent, text="Project Text Source").pack(anchor="w", padx=8)
        self.project_text = tk.Text(parent, height=20, wrap="none", font=("Consolas", 11))
        self.project_text.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        text_btns = ttk.Frame(parent)
        text_btns.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(text_btns, text="Rebuild From Text", command=self.rebuild_project_from_text).pack(side="left")
        ttk.Button(text_btns, text="Refresh Text", command=lambda: self.refresh_text_from_project(mark_clean=False)).pack(side="left", padx=6)

        ttk.Label(parent, text="Timeline").pack(anchor="w", padx=8)
        timeline_frame = ttk.Frame(parent)
        timeline_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.timeline = TimelineCanvas(timeline_frame, self, height=260)
        x_scroll = ttk.Scrollbar(timeline_frame, orient="horizontal", command=self.timeline.xview)
        y_scroll = ttk.Scrollbar(timeline_frame, orient="vertical", command=self.timeline.yview)
        self.timeline.configure(xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)
        self.timeline.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        timeline_frame.rowconfigure(0, weight=1)
        timeline_frame.columnconfigure(0, weight=1)

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

    def get_slide(self, slide_id: Optional[str]) -> Optional[Slide]:
        if not slide_id:
            return None
        for slide in self.project.slides:
            if slide.slide_id == slide_id:
                return slide
        return None

    def refresh_slide_list(self) -> None:
        self.project.sort_slides()
        self.project.normalize_ids()
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
        self.refresh_editor_fields_from_selected(preserve_text_focus=False)
        self.timeline.redraw()
        self.render_preview()

    def refresh_editor_fields_from_selected(self, preserve_text_focus: bool = True) -> None:
        slide = self.get_slide(self.selected_slide_id)
        self._updating_ui = True
        try:
            if slide is None:
                self.title_var.set("")
                self.start_var.set("")
                self.duration_var.set("")
                self.layer_var.set("")
                self.background_var.set("")
                self.crop_x_var.set("")
                self.crop_y_var.set("")
                self.body_text.delete("1.0", tk.END)
                self.narration_text.delete("1.0", tk.END)
            else:
                self.title_var.set(slide.title)
                self.start_var.set(f"{slide.start:.2f}")
                self.duration_var.set(f"{slide.duration:.2f}")
                self.layer_var.set(str(slide.layer))
                self.background_var.set(slide.background)
                self.crop_x_var.set(f"{slide.crop_x:.3f}")
                self.crop_y_var.set(f"{slide.crop_y:.3f}")
                self.body_text.delete("1.0", tk.END)
                self.body_text.insert("1.0", slide.body)
                self.narration_text.delete("1.0", tk.END)
                self.narration_text.insert("1.0", slide.narration)
        finally:
            self._updating_ui = False

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
        except ValueError:
            messagebox.showerror("Invalid value", "Start, duration, layer, and crop fields must contain valid numbers.")
            return
        self.project.sort_slides()
        self.refresh_slide_list()
        self.timeline.redraw()
        self.render_preview()
        self.refresh_text_from_project(mark_clean=False)
        self.mark_dirty()
        self.set_status("Slide updated")

    def focus_title_entry(self) -> None:
        self.title_entry.focus_set()
        self.title_entry.selection_range(0, tk.END)

    def on_slide_list_select(self, event=None) -> None:
        selection = self.slide_list.curselection()
        if not selection:
            return
        idx = selection[0]
        if 0 <= idx < len(self.project.slides):
            self.select_slide(self.project.slides[idx].slide_id)

    def on_field_change(self, *args) -> None:
        if self._updating_ui:
            return
        self.mark_dirty()

    def on_text_editor_change(self, event=None) -> None:
        if self._updating_ui:
            return
        slide = self.get_slide(self.selected_slide_id)
        if slide is None:
            return
        slide.body = self.body_text.get("1.0", tk.END).strip()
        slide.narration = self.narration_text.get("1.0", tk.END).strip()
        self.refresh_text_from_project(mark_clean=False)
        self.render_preview()
        self.mark_dirty()

    def add_slide(self) -> None:
        self.project.normalize_ids()
        next_idx = len(self.project.slides) + 1
        start = self.project.total_duration()
        slide = Slide(
            slide_id=f"slide_{next_idx:03d}",
            title=f"Slide {next_idx}",
            start=round(start, 2),
            duration=5.0,
            layer=0,
            body="Add body text here.",
            narration="Add narration here.",
        )
        self.project.slides.append(slide)
        self.project.sort_slides()
        self.select_slide(slide.slide_id)
        self.refresh_text_from_project(mark_clean=False)
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
        self.refresh_text_from_project(mark_clean=False)
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
        new_sel = self.project.slides[0].slide_id if self.project.slides else None
        self.select_slide(new_sel)
        self.refresh_text_from_project(mark_clean=False)
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
        self.refresh_text_from_project(mark_clean=False)
        self.mark_dirty()
        self.set_status("Slides auto-placed sequentially")

    def import_background_image(self) -> None:
        slide = self.get_slide(self.selected_slide_id)
        if slide is None:
            return
        path = filedialog.askopenfilename(
            title="Select background image",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.webp *.bmp"), ("All files", "*.*")],
        )
        if not path:
            return
        slide.background = path
        self.background_var.set(path)
        self.refresh_text_from_project(mark_clean=False)
        self.render_preview()
        self.mark_dirty()
        self.set_status("Background image linked to slide")

    def refresh_text_from_project(self, mark_clean: bool = False) -> None:
        text = project_to_text(self.project)
        cursor = self.project_text.index(tk.INSERT)
        self.project_text.delete("1.0", tk.END)
        self.project_text.insert("1.0", text)
        try:
            self.project_text.mark_set(tk.INSERT, cursor)
        except Exception:
            pass
        if mark_clean:
            self.mark_clean()

    def rebuild_project_from_text(self) -> None:
        text = self.project_text.get("1.0", tk.END)
        try:
            new_project = parse_project_text(text)
        except Exception as exc:
            messagebox.showerror("Parse error", f"Could not rebuild project from text.\n\n{exc}")
            return
        self.project = new_project
        self.selected_slide_id = self.project.slides[0].slide_id if self.project.slides else None
        self.refresh_slide_list()
        self.refresh_editor_fields_from_selected(preserve_text_focus=False)
        self.timeline.redraw()
        self.render_preview()
        self.mark_dirty()
        self.set_status("Project rebuilt from text")

    def export_plain_text(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Export plain text project",
            defaultextension=".vsl.txt",
            filetypes=[("VSL Project", "*.vsl.txt"), ("Text files", "*.txt"), ("All files", "*.*")],
        )
        if not path:
            return
        Path(path).write_text(project_to_text(self.project), encoding="utf-8")
        self.set_status(f"Exported plain text: {path}")

    def new_project(self) -> None:
        if not self.confirm_discard_changes():
            return
        self.project = Project(slides=[])
        self.current_path = None
        self.selected_slide_id = None
        self.refresh_slide_list()
        self.refresh_editor_fields_from_selected(preserve_text_focus=False)
        self.refresh_text_from_project(mark_clean=True)
        self.timeline.redraw()
        self.render_preview()
        self.set_status("New project created")

    def open_project(self) -> None:
        if not self.confirm_discard_changes():
            return
        path = filedialog.askopenfilename(
            title="Open VSL project",
            filetypes=[("VSL Project", "*.vsl.txt *.txt"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            text = Path(path).read_text(encoding="utf-8")
            self.project = parse_project_text(text)
            self.current_path = Path(path)
            self.selected_slide_id = self.project.slides[0].slide_id if self.project.slides else None
            self.refresh_slide_list()
            self.refresh_editor_fields_from_selected(preserve_text_focus=False)
            self.refresh_text_from_project(mark_clean=True)
            self.timeline.redraw()
            self.render_preview()
            self.set_status(f"Opened: {path}")
        except Exception as exc:
            messagebox.showerror("Open failed", f"Could not open file.\n\n{exc}")

    def save_project(self) -> None:
        if self.current_path is None:
            self.save_project_as()
            return
        try:
            self.apply_editor_to_slide_if_possible(silent=True)
            self.current_path.write_text(project_to_text(self.project), encoding="utf-8")
            self.mark_clean()
            self.set_status(f"Saved: {self.current_path}")
        except Exception as exc:
            messagebox.showerror("Save failed", f"Could not save file.\n\n{exc}")

    def save_project_as(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save VSL project",
            defaultextension=".vsl.txt",
            filetypes=[("VSL Project", "*.vsl.txt"), ("Text files", "*.txt"), ("All files", "*.*")],
        )
        if not path:
            return
        self.current_path = Path(path)
        self.save_project()

    def apply_editor_to_slide_if_possible(self, silent: bool = False) -> None:
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
            self.project.sort_slides()
        except ValueError:
            if not silent:
                raise

    def confirm_discard_changes(self) -> bool:
        if not self.is_dirty:
            return True
        return messagebox.askyesno("Unsaved changes", "Discard unsaved changes?")

    def render_preview(self) -> None:
        self.preview_canvas.delete("all")
        slide = self.get_slide(self.selected_slide_id)
        canvas_w = max(1, self.preview_canvas.winfo_width())
        canvas_h = max(1, self.preview_canvas.winfo_height())
        self.preview_canvas.create_rectangle(0, 0, canvas_w, canvas_h, fill=PREVIEW_BG, outline="")

        if slide is None:
            self.preview_canvas.create_text(canvas_w / 2, canvas_h / 2, text="No slide selected", fill="#cbd5e1", font=TITLE_FONT)
            return

        if PIL_AVAILABLE and slide.background and os.path.exists(slide.background):
            try:
                img = Image.open(slide.background).convert("RGB")
                fitted = ImageOps.fit(img, (canvas_w, canvas_h), centering=(slide.crop_x, slide.crop_y))
                self.preview_image_cache = ImageTk.PhotoImage(fitted)
                self.preview_canvas.create_image(0, 0, image=self.preview_image_cache, anchor="nw")
                self.preview_canvas.create_rectangle(0, 0, canvas_w, canvas_h, fill="#000000", outline="", stipple="gray50")
            except Exception:
                self.preview_canvas.create_rectangle(0, 0, canvas_w, canvas_h, fill="#1f2937", outline="")
        else:
            self.preview_canvas.create_rectangle(0, 0, canvas_w, canvas_h, fill="#1f2937", outline="")

        self.preview_canvas.create_rectangle(18, 18, canvas_w - 18, canvas_h - 18, outline="#475569")
        self.preview_canvas.create_text(40, 36, anchor="nw", text=slide.title, fill="#ffffff", font=TITLE_FONT, width=canvas_w - 80)

        body_text = slide.body.strip() or "Body text goes here."
        wrapped = textwrap.fill(body_text, width=max(20, int(canvas_w / 12)))
        self.preview_canvas.create_text(40, 100, anchor="nw", text=wrapped, fill="#e2e8f0", font=BODY_FONT, width=canvas_w - 80)

        narration_text = slide.narration.strip() or "Narration text goes here."
        narration_preview = textwrap.shorten(narration_text.replace("\n", " "), width=180, placeholder=" ...")
        self.preview_canvas.create_rectangle(30, canvas_h - 90, canvas_w - 30, canvas_h - 30, fill="#0f172a", outline="#334155")
        self.preview_canvas.create_text(40, canvas_h - 80, anchor="nw", text=f"Narration: {narration_preview}", fill="#f8fafc", font=("Arial", 11), width=canvas_w - 80)
        self.preview_canvas.create_text(canvas_w - 12, 10, anchor="ne", text=f"{slide.start:.2f}s • {slide.duration:.2f}s • L{slide.layer}", fill="#cbd5e1", font=("Arial", 10))

        if not PIL_AVAILABLE:
            self.preview_canvas.create_text(canvas_w - 12, canvas_h - 10, anchor="se", text="Install pillow for image crop preview", fill="#94a3b8", font=("Arial", 9))

    def on_preview_crop_press(self, event) -> None:
        slide = self.get_slide(self.selected_slide_id)
        if slide is None:
            return
        self._drag_crop = True
        self.update_crop_from_preview(event.x, event.y)

    def on_preview_crop_drag(self, event) -> None:
        if not self._drag_crop:
            return
        self.update_crop_from_preview(event.x, event.y)

    def on_preview_crop_release(self, event) -> None:
        if not self._drag_crop:
            return
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
        self.refresh_text_from_project(mark_clean=False)
        self.set_status("Crop center updated from preview drag")


def main() -> None:
    root = tk.Tk()
    try:
        style = ttk.Style(root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass
    app = VSLFoundationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
