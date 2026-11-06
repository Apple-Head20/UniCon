"""
OmniConvert Metro — Universal File Converter
Windows 8 Metro UI · All-format conversion · Animated tiles
Dependencies:
    pip install customtkinter tkinterdnd2 Pillow pillow-heif moviepy
    LibreOffice must be installed for document conversion
    ffmpeg must be in PATH for video/audio conversion
"""

import os
import sys
import subprocess
import tempfile
import shutil
import concurrent.futures
import threading
import re
import math
import time
import colorsys
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

# ── Optional deps ──────────────────────────────────────────────────────────────
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_OK = True
except ImportError:
    DND_OK = False

try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
    HEIF_OK = True
except ImportError:
    HEIF_OK = False

try:
    from PIL import Image, ImageTk, ImageDraw, ImageFilter
    PIL_OK = True
except ImportError:
    PIL_OK = False

# ══════════════════════════════════════════════════════════════════════════════
#  METRO PALETTE & THEME
# ══════════════════════════════════════════════════════════════════════════════
ctk.set_appearance_mode("dark")

C = {
    "bg":           "#0f0f0f",
    "panel":        "#1a1a1a",
    "tile":         "#1e1e1e",
    "border":       "#2a2a2a",
    "accent":       "#0078d7",
    "accent2":      "#00b4d8",
    "accent3":      "#e74856",
    "accent4":      "#16c60c",
    "accent5":      "#ca5010",
    "accent6":      "#8764b8",
    "accent7":      "#038387",
    "accent8":      "#c239b3",
    "text":         "#ffffff",
    "text_dim":     "#888888",
    "text_darker":  "#444444",
    "hover":        "#262626",
    "selected":     "#003f73",
}

TILE_COLORS = [
    C["accent"],  C["accent2"], C["accent3"], C["accent4"],
    C["accent5"], C["accent6"], C["accent7"], C["accent8"],
]

# Font choices  (fall back gracefully)
FONT_DISPLAY = "Segoe UI Light"
FONT_BODY    = "Segoe UI"
FONT_MONO    = "Consolas"

def F(size, weight="normal", family=None):
    fam = family or (FONT_DISPLAY if size >= 18 else FONT_BODY)
    return ctk.CTkFont(family=fam, size=size, weight=weight)

# ══════════════════════════════════════════════════════════════════════════════
#  CONVERSION MAP  –  (input-extensions) → {label, output-formats, engine}
# ══════════════════════════════════════════════════════════════════════════════
CONV_MAP = {
    # ── Documents ──────────────────────────────────────────────────────────
    "Documents": {
        "exts": [".docx",".doc",".odt",".rtf",".txt",".pptx",".ppt",".odp",
                 ".xlsx",".xls",".ods",".csv"],
        "outputs": ["PDF","DOCX","ODT","RTF","TXT","PPTX","ODP","XLSX","ODS","HTML","CSV"],
        "engine": "libreoffice",
        "icon": "📄",
        "color": C["accent"],
    },
    # ── Images ─────────────────────────────────────────────────────────────
    "Images": {
        "exts": [".jpg",".jpeg",".png",".webp",".bmp",".tiff",".tif",
                 ".gif",".ico",".heic",".heif",".avif",".svg",".psd"],
        "outputs": ["JPG","PNG","WEBP","BMP","TIFF","GIF","ICO","PDF","AVIF"],
        "engine": "pillow",
        "icon": "🖼️",
        "color": C["accent2"],
    },
    # ── Video ──────────────────────────────────────────────────────────────
    "Video": {
        "exts": [".mp4",".mkv",".avi",".mov",".wmv",".flv",".webm",
                 ".m4v",".3gp",".ts",".vob",".ogv"],
        "outputs": ["MP4","MKV","AVI","MOV","WEBM","GIF","MP3","AAC","WAV","FLV"],
        "engine": "ffmpeg",
        "icon": "🎬",
        "color": C["accent3"],
    },
    # ── Audio ──────────────────────────────────────────────────────────────
    "Audio": {
        "exts": [".mp3",".wav",".flac",".aac",".ogg",".m4a",
                 ".wma",".opus",".aiff",".alac"],
        "outputs": ["MP3","WAV","FLAC","AAC","OGG","M4A","OPUS","AIFF"],
        "engine": "ffmpeg",
        "icon": "🎵",
        "color": C["accent4"],
    },
    # ── Archives ───────────────────────────────────────────────────────────
    "Archives": {
        "exts": [".zip",".tar",".gz",".bz2",".xz",".7z",".rar"],
        "outputs": ["ZIP","TAR","TAR.GZ","TAR.BZ2","TAR.XZ"],
        "engine": "archive",
        "icon": "🗜️",
        "color": C["accent5"],
    },
    # ── Ebooks ─────────────────────────────────────────────────────────────
    "Ebooks": {
        "exts": [".epub",".mobi",".azw3",".fb2",".pdf"],
        "outputs": ["EPUB","MOBI","AZW3","PDF","TXT"],
        "engine": "calibre",
        "icon": "📚",
        "color": C["accent6"],
    },
    # ── Data ───────────────────────────────────────────────────────────────
    "Data": {
        "exts": [".csv",".json",".xml",".yaml",".yml",".tsv",".xlsx",".sql"],
        "outputs": ["CSV","JSON","XML","YAML","TSV","XLSX","SQL"],
        "engine": "data",
        "icon": "🗃️",
        "color": C["accent7"],
    },
    # ── Fonts ──────────────────────────────────────────────────────────────
    "Fonts": {
        "exts": [".ttf",".otf",".woff",".woff2"],
        "outputs": ["TTF","OTF","WOFF","WOFF2"],
        "engine": "fonttools",
        "icon": "🔤",
        "color": C["accent8"],
    },
}

ALL_EXTS = {}
for cat, info in CONV_MAP.items():
    for ext in info["exts"]:
        ALL_EXTS[ext] = cat

# ══════════════════════════════════════════════════════════════════════════════
#  HELPER: detect tools
# ══════════════════════════════════════════════════════════════════════════════
def _which(cmd):
    return shutil.which(cmd) is not None

def get_soffice():
    if sys.platform == "win32":
        for p in [r"C:\Program Files\LibreOffice\program\soffice.exe",
                  r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]:
            if os.path.exists(p): return p
    return "libreoffice"

SOFFICE   = get_soffice()
HAS_FF    = _which("ffmpeg")
HAS_CAL   = _which("ebook-convert")   # Calibre
HAS_FONTS = False
try:
    import fonttools  # noqa
    HAS_FONTS = True
except ImportError:
    pass

# ══════════════════════════════════════════════════════════════════════════════
#  ANIMATED CANVAS BACKGROUND  (Metro live-tile rain effect)
# ══════════════════════════════════════════════════════════════════════════════
class MetroBackground(tk.Canvas):
    """Subtle animated squares drifting upward – Metro live-tile feel."""
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=C["bg"], highlightthickness=0, **kw)
        self._particles = []
        self._running   = True
        self.bind("<Configure>", self._on_resize)
        self._spawn_initial()
        self._tick()

    def _on_resize(self, e):
        self._spawn_initial()

    def _rand(self, lo, hi):
        import random
        return random.uniform(lo, hi)

    def _spawn_initial(self):
        import random
        self._particles.clear()
        self.delete("particle")
        w = self.winfo_width()  or 960
        h = self.winfo_height() or 700
        for _ in range(18):
            self._spawn(random.uniform(0, w), random.uniform(0, h))

    def _spawn(self, x=None, y=None):
        import random
        w = self.winfo_width()  or 960
        h = self.winfo_height() or 700
        x = x if x is not None else random.uniform(0, w)
        y = y if y is not None else h + 20
        size   = random.randint(18, 70)
        speed  = random.uniform(0.25, 0.8)
        col    = random.choice(TILE_COLORS)
        alpha_hex = "%02x" % random.randint(12, 35)
        # Tkinter canvas doesn't support rgba so we fake alpha with close-to-bg colour
        r0,g0,b0 = int(C["bg"][1:3],16), int(C["bg"][3:5],16), int(C["bg"][5:7],16)
        r1,g1,b1 = int(col[1:3],16), int(col[3:5],16), int(col[5:7],16)
        t = random.uniform(0.05, 0.18)
        mixed = "#%02x%02x%02x" % (
            int(r0+(r1-r0)*t), int(g0+(g1-g0)*t), int(b0+(b1-b0)*t))
        rid = self.create_rectangle(x, y, x+size, y+size,
                                    fill=mixed, outline="", tags="particle")
        self._particles.append({"id": rid, "x": x, "y": y,
                                 "size": size, "speed": speed, "col": col})

    def _tick(self):
        if not self._running:
            return
        import random
        w = self.winfo_width() or 960
        dead = []
        for p in self._particles:
            p["y"] -= p["speed"]
            self.coords(p["id"], p["x"], p["y"],
                        p["x"]+p["size"], p["y"]+p["size"])
            if p["y"] + p["size"] < 0:
                dead.append(p)
        for p in dead:
            self.delete(p["id"])
            self._particles.remove(p)
            self._spawn()
        self.after(30, self._tick)

    def stop(self):
        self._running = False


# ══════════════════════════════════════════════════════════════════════════════
#  METRO TILE BUTTON
# ══════════════════════════════════════════════════════════════════════════════
class MetroTile(ctk.CTkFrame):
    """Clickable Metro-style tile with icon, label, and hover animation."""
    def __init__(self, parent, icon, label, color, size=(140,100), command=None, **kw):
        super().__init__(parent, width=size[0], height=size[1],
                         corner_radius=0, fg_color=color,
                         border_width=0, **kw)
        self.pack_propagate(False)
        self._color     = color
        self._hov_color = self._lighten(color, 0.15)
        self._command   = command
        self._anim_id   = None
        self._current   = color

        self._icon_lbl = ctk.CTkLabel(self, text=icon, font=F(28), fg_color="transparent",
                                       text_color=C["text"])
        self._icon_lbl.pack(pady=(14, 2))
        self._text_lbl = ctk.CTkLabel(self, text=label.upper(),
                                       font=F(11, "bold"),
                                       fg_color="transparent", text_color=C["text"])
        self._text_lbl.pack()

        for w in (self, self._icon_lbl, self._text_lbl):
            w.bind("<Enter>",    self._on_enter)
            w.bind("<Leave>",    self._on_leave)
            w.bind("<Button-1>", self._on_click)

    def _lighten(self, hex_color, amount):
        r,g,b = [int(hex_color[i:i+2],16)/255 for i in (1,3,5)]
        h,l,s = colorsys.rgb_to_hls(r,g,b)
        l = min(1, l + amount)
        r,g,b = [int(x*255) for x in colorsys.hls_to_rgb(h,l,s)]
        return "#%02x%02x%02x" % (r,g,b)

    def _on_enter(self, e):
        self._animate_to(self._hov_color)

    def _on_leave(self, e):
        self._animate_to(self._color)

    def _on_click(self, e):
        if self._command:
            self._command()

    def _animate_to(self, target):
        if self._anim_id:
            self.after_cancel(self._anim_id)
        self._target = target
        self._step_color()

    def _step_color(self):
        src = self._current
        tgt = self._target
        if src == tgt:
            return
        def interp(a, b, t=0.25):
            return int(a + (b-a)*t)
        r = interp(int(src[1:3],16), int(tgt[1:3],16))
        g = interp(int(src[3:5],16), int(tgt[3:5],16))
        b = interp(int(src[5:7],16), int(tgt[5:7],16))
        new = "#%02x%02x%02x" % (r,g,b)
        self._current = new
        self.configure(fg_color=new)
        if abs(int(src[1:3],16)-int(tgt[1:3],16)) > 2 or \
           abs(int(src[3:5],16)-int(tgt[3:5],16)) > 2 or \
           abs(int(src[5:7],16)-int(tgt[5:7],16)) > 2:
            self._anim_id = self.after(16, self._step_color)


# ══════════════════════════════════════════════════════════════════════════════
#  METRO PROGRESS BAR (animated shimmer)
# ══════════════════════════════════════════════════════════════════════════════
class MetroProgressBar(tk.Canvas):
    def __init__(self, parent, color=C["accent"], height=6, **kw):
        super().__init__(parent, bg=C["panel"], highlightthickness=0,
                         height=height, **kw)
        self._color   = color
        self._value   = 0.0
        self._shimmer = 0
        self._running = False
        self._bar_id  = None
        self._shim_id = None
        self.bind("<Configure>", self._redraw)

    def set(self, v):
        self._value = max(0.0, min(1.0, v))
        self._redraw()

    def _redraw(self, e=None):
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 2: return
        # track
        self.create_rectangle(0,0,w,h, fill=C["border"], outline="")
        # fill
        fill_w = int(w * self._value)
        if fill_w > 0:
            self._bar_id = self.create_rectangle(0,0,fill_w,h,
                                                  fill=self._color, outline="")
            # shimmer stripe
            s = self._shimmer % (fill_w + 60) - 30
            self.create_rectangle(s, 0, s+20, h,
                                   fill=self._lighten(self._color, 0.3),
                                   outline="", stipple="gray50")

    def _lighten(self, hex_color, amount):
        r,g,b = [int(hex_color[i:i+2],16)/255 for i in (1,3,5)]
        h,l,s = colorsys.rgb_to_hls(r,g,b)
        l = min(1, l + amount)
        r,g,b = [int(x*255) for x in colorsys.hls_to_rgb(h,l,s)]
        return "#%02x%02x%02x" % (r,g,b)

    def start_shimmer(self):
        self._running = True
        self._animate()

    def stop_shimmer(self):
        self._running = False

    def _animate(self):
        if not self._running: return
        self._shimmer += 4
        self._redraw()
        self.after(30, self._animate)


# ══════════════════════════════════════════════════════════════════════════════
#  QUEUE ROW
# ══════════════════════════════════════════════════════════════════════════════
class QueueRow(ctk.CTkFrame):
    STATES = {"waiting": C["text_dim"], "converting": C["accent2"],
              "done": C["accent4"], "error": C["accent3"]}

    def __init__(self, parent, filepath, on_remove, **kw):
        super().__init__(parent, corner_radius=0,
                         fg_color=C["tile"], height=42, **kw)
        self.pack_propagate(False)
        self.filepath = filepath
        name = os.path.basename(filepath)
        ext  = os.path.splitext(name)[1].lower()
        cat  = ALL_EXTS.get(ext, "?")
        icon = CONV_MAP[cat]["icon"] if cat in CONV_MAP else "📁"

        # icon
        ctk.CTkLabel(self, text=icon, font=F(14), fg_color="transparent",
                     width=30).pack(side="left", padx=(8,4))
        # filename
        self._name_lbl = ctk.CTkLabel(self, text=name, font=F(12),
                                       fg_color="transparent",
                                       text_color=C["text"], anchor="w")
        self._name_lbl.pack(side="left", fill="x", expand=True, padx=4)
        # status dot
        self._status = ctk.CTkLabel(self, text="●", font=F(10),
                                     text_color=C["text_dim"],
                                     fg_color="transparent", width=20)
        self._status.pack(side="right", padx=4)
        # remove
        ctk.CTkButton(self, text="✕", width=28, height=28,
                      corner_radius=0, fg_color="transparent",
                      hover_color=C["accent3"], font=F(11,"bold"),
                      text_color=C["text_dim"],
                      command=lambda: on_remove(filepath)
                      ).pack(side="right", padx=4)

        # hover flash
        self.bind("<Enter>", lambda e: self.configure(fg_color=C["hover"]))
        self.bind("<Leave>", lambda e: self.configure(fg_color=C["tile"]))

    def set_state(self, state):
        col  = self.STATES.get(state, C["text_dim"])
        syms = {"waiting":"●","converting":"◉","done":"✔","error":"✘"}
        self._status.configure(text=syms.get(state,"●"), text_color=col)
        if state == "converting":
            self._name_lbl.configure(text_color=C["accent2"])
        elif state == "done":
            self._name_lbl.configure(text_color=C["accent4"])
        elif state == "error":
            self._name_lbl.configure(text_color=C["accent3"])


# ══════════════════════════════════════════════════════════════════════════════
#  CONVERSION ENGINE
# ══════════════════════════════════════════════════════════════════════════════
class ConversionEngine:
    @staticmethod
    def detect_category(filepath):
        ext = os.path.splitext(filepath)[1].lower()
        return ALL_EXTS.get(ext)

    # ── LibreOffice ───────────────────────────────────────────────────────
    @staticmethod
    def convert_libreoffice(input_file, out_dir, target_fmt):
        tmp = tempfile.mkdtemp()
        try:
            out_unix = out_dir.replace("\\", "/")
            tmp_unix = tmp.replace("\\", "/")
            cmd = [SOFFICE, "--headless",
                   f"-env:UserInstallation=file:///{tmp_unix}",
                   "--convert-to", target_fmt.lower(),
                   "--outdir", out_unix, input_file]
            flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            subprocess.run(cmd, check=True,
                           stdout=subprocess.DEVNULL,
                           stderr=subprocess.DEVNULL,
                           creationflags=flags)
            return True
        except Exception as e:
            print(f"LibreOffice error: {e}")
            return False
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    # ── Pillow ────────────────────────────────────────────────────────────
    @staticmethod
    def convert_pillow(input_file, out_dir, target_fmt):
        try:
            img  = Image.open(input_file)
            stem = os.path.splitext(os.path.basename(input_file))[0]
            save_fmt = {"JPG":"JPEG","AVIF":"AVIF"}.get(target_fmt, target_fmt)
            # transparency fix
            if save_fmt in ("JPEG","BMP","ICO") and img.mode in ("RGBA","P","LA"):
                bg = Image.new("RGB", img.size, (255,255,255))
                bg.paste(img, mask=img.split()[-1] if img.mode in ("RGBA","LA") else None)
                img = bg
            elif save_fmt == "ICO":
                img = img.resize((256,256), Image.LANCZOS)
            out_path = os.path.join(out_dir,
                                    f"{stem}.{target_fmt.lower().replace('jpg','jpg')}")
            img.save(out_path, format=save_fmt)
            return True
        except Exception as e:
            print(f"Pillow error: {e}")
            return False

    # ── FFmpeg ────────────────────────────────────────────────────────────
    @staticmethod
    def convert_ffmpeg(input_file, out_dir, target_fmt):
        if not HAS_FF:
            return False
        stem = os.path.splitext(os.path.basename(input_file))[0]
        ext  = target_fmt.lower()
        out  = os.path.join(out_dir, f"{stem}.{ext}")
        cmd  = ["ffmpeg", "-y", "-i", input_file]
        # audio-only extract from video
        if ext in ("mp3","aac","wav","flac","ogg","opus","m4a","aiff"):
            cmd += ["-vn"]
        # gif special
        if ext == "gif":
            cmd += ["-vf", "fps=12,scale=480:-1:flags=lanczos",
                    "-loop", "0"]
        # quality presets
        if ext == "mp4":
            cmd += ["-c:v","libx264","-crf","23","-c:a","aac","-b:a","192k"]
        elif ext == "webm":
            cmd += ["-c:v","libvpx-vp9","-crf","33","-b:v","0","-c:a","libopus"]
        elif ext == "mkv":
            cmd += ["-c:v","libx264","-crf","23","-c:a","copy"]
        cmd.append(out)
        flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        try:
            subprocess.run(cmd, check=True,
                           stdout=subprocess.DEVNULL,
                           stderr=subprocess.DEVNULL,
                           creationflags=flags)
            return True
        except Exception as e:
            print(f"FFmpeg error: {e}")
            return False

    # ── Archive ───────────────────────────────────────────────────────────
    @staticmethod
    def convert_archive(input_file, out_dir, target_fmt):
        import zipfile, tarfile
        stem    = os.path.splitext(os.path.basename(input_file))[0]
        fmt     = target_fmt.upper()
        out     = os.path.join(out_dir, stem)
        tmp     = tempfile.mkdtemp()
        try:
            # Extract first
            src_ext = os.path.splitext(input_file)[1].lower()
            if src_ext == ".zip":
                with zipfile.ZipFile(input_file) as z:
                    z.extractall(tmp)
            elif src_ext in (".tar",".gz",".bz2",".xz"):
                with tarfile.open(input_file) as t:
                    t.extractall(tmp)
            else:
                # try 7z / unrar via subprocess
                subprocess.run(["7z","x",input_file,f"-o{tmp}","-y"],
                               check=True,
                               stdout=subprocess.DEVNULL,
                               stderr=subprocess.DEVNULL)
            # Re-pack
            if fmt == "ZIP":
                with zipfile.ZipFile(out+".zip","w",
                                     zipfile.ZIP_DEFLATED) as z:
                    for root, dirs, files in os.walk(tmp):
                        for f in files:
                            fp = os.path.join(root, f)
                            z.write(fp, os.path.relpath(fp, tmp))
            elif fmt in ("TAR","TAR.GZ","TAR.BZ2","TAR.XZ"):
                mode_map = {"TAR":"w","TAR.GZ":"w:gz",
                            "TAR.BZ2":"w:bz2","TAR.XZ":"w:xz"}
                ext_map  = {"TAR":".tar","TAR.GZ":".tar.gz",
                            "TAR.BZ2":".tar.bz2","TAR.XZ":".tar.xz"}
                with tarfile.open(out+ext_map[fmt], mode_map[fmt]) as t:
                    t.add(tmp, arcname="")
            return True
        except Exception as e:
            print(f"Archive error: {e}")
            return False
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    # ── Calibre ────────────────────────────────────────────────────────────
    @staticmethod
    def convert_calibre(input_file, out_dir, target_fmt):
        if not HAS_CAL:
            return False
        stem = os.path.splitext(os.path.basename(input_file))[0]
        out  = os.path.join(out_dir, f"{stem}.{target_fmt.lower()}")
        try:
            subprocess.run(["ebook-convert", input_file, out],
                           check=True,
                           stdout=subprocess.DEVNULL,
                           stderr=subprocess.DEVNULL)
            return True
        except Exception as e:
            print(f"Calibre error: {e}")
            return False

    # ── Data formats ───────────────────────────────────────────────────────
    @staticmethod
    def convert_data(input_file, out_dir, target_fmt):
        try:
            import json, csv, io
            stem = os.path.splitext(os.path.basename(input_file))[0]
            out  = os.path.join(out_dir, f"{stem}.{target_fmt.lower()}")
            src_ext = os.path.splitext(input_file)[1].lower()

            # Load source
            data = None
            with open(input_file, "r", encoding="utf-8", errors="replace") as f:
                raw = f.read()
            if src_ext == ".json":
                data = json.loads(raw)
            elif src_ext in (".csv",".tsv"):
                delim = "\t" if src_ext == ".tsv" else ","
                reader = csv.DictReader(io.StringIO(raw), delimiter=delim)
                data = list(reader)
            elif src_ext in (".yaml",".yml"):
                import yaml
                data = yaml.safe_load(raw)
            elif src_ext == ".xml":
                import xml.etree.ElementTree as ET
                root = ET.fromstring(raw)
                data = {root.tag: ConversionEngine._xml_to_dict(root)}
            else:
                data = raw  # fallback: treat as string

            fmt = target_fmt.upper()
            with open(out, "w", encoding="utf-8", newline="") as f:
                if fmt == "JSON":
                    json.dump(data if isinstance(data, (dict,list)) else {"data":data},
                              f, indent=2)
                elif fmt in ("CSV","TSV"):
                    delim = "\t" if fmt == "TSV" else ","
                    if isinstance(data, list) and data:
                        w = csv.DictWriter(f, fieldnames=data[0].keys(),
                                           delimiter=delim)
                        w.writeheader(); w.writerows(data)
                    else:
                        f.write(str(data))
                elif fmt in ("YAML","YML"):
                    import yaml
                    yaml.dump(data, f, allow_unicode=True)
                elif fmt == "XML":
                    f.write(ConversionEngine._dict_to_xml(
                        data if isinstance(data, dict) else {"root": data}))
                else:
                    f.write(str(data))
            return True
        except Exception as e:
            print(f"Data conv error: {e}")
            return False

    @staticmethod
    def _xml_to_dict(el):
        d = {}
        for child in el:
            d[child.tag] = ConversionEngine._xml_to_dict(child) or child.text
        return d or el.text

    @staticmethod
    def _dict_to_xml(d, root="root"):
        def _recurse(d, name):
            el = f"<{name}>"
            if isinstance(d, dict):
                for k,v in d.items():
                    el += _recurse(v, k)
            elif isinstance(d, list):
                for item in d:
                    el += _recurse(item, "item")
            else:
                el += str(d) if d is not None else ""
            el += f"</{name}>"
            return el
        return '<?xml version="1.0" encoding="UTF-8"?>\n' + _recurse(d, root)

    # ── Font tools ─────────────────────────────────────────────────────────
    @staticmethod
    def convert_font(input_file, out_dir, target_fmt):
        try:
            from fontTools.ttLib import TTFont
            stem = os.path.splitext(os.path.basename(input_file))[0]
            out  = os.path.join(out_dir, f"{stem}.{target_fmt.lower()}")
            font = TTFont(input_file)
            font.save(out)
            return True
        except Exception as e:
            print(f"Font error: {e}")
            return False

    # ── DISPATCHER ─────────────────────────────────────────────────────────
    @classmethod
    def convert(cls, filepath, out_dir, target_fmt, engine):
        if engine == "libreoffice":
            return cls.convert_libreoffice(filepath, out_dir, target_fmt)
        elif engine == "pillow":
            return cls.convert_pillow(filepath, out_dir, target_fmt)
        elif engine == "ffmpeg":
            return cls.convert_ffmpeg(filepath, out_dir, target_fmt)
        elif engine == "archive":
            return cls.convert_archive(filepath, out_dir, target_fmt)
        elif engine == "calibre":
            return cls.convert_calibre(filepath, out_dir, target_fmt)
        elif engine == "data":
            return cls.convert_data(filepath, out_dir, target_fmt)
        elif engine == "fonttools":
            return cls.convert_font(filepath, out_dir, target_fmt)
        return False


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ══════════════════════════════════════════════════════════════════════════════
Base = (ctk.CTk if not DND_OK
        else type("_Base", (ctk.CTk, TkinterDnD.DnDWrapper), {}))

class OmniConvertApp(Base):
    def __init__(self):
        super().__init__()
        if DND_OK:
            self.TkdndVersion = TkinterDnD._require(self)

        self.title("OmniConvert Metro")
        self.geometry("980x720")
        self.minsize(820, 600)
        self.configure(fg_color=C["bg"])

        self._files    = []        # list of filepath strings
        self._rows     = {}        # filepath → QueueRow
        self._out_dir  = ""
        self._busy     = False

        self._build_ui()
        self._check_deps()

    # ── UI Construction ───────────────────────────────────────────────────

    def _build_ui(self):
        # Animated background (placed behind everything)
        self._bg_canvas = MetroBackground(self, width=980, height=720)
        self._bg_canvas.place(x=0, y=0, relwidth=1, relheight=1)

        # ── TOP HEADER ──────────────────────────────────────────────────
        header = tk.Frame(self, bg=C["bg"], height=70)
        header.pack(fill="x", padx=0, pady=0)
        header.pack_propagate(False)

        # App title — big Metro typography
        tk.Label(header, text="OMNICONVERT", font=(FONT_DISPLAY, 28),
                 fg=C["text"], bg=C["bg"]).pack(side="left", padx=24, pady=14)
        tk.Label(header, text="METRO", font=(FONT_DISPLAY, 28),
                 fg=C["accent"], bg=C["bg"]).pack(side="left", pady=14)
        tk.Label(header, text="  universal file converter",
                 font=(FONT_BODY, 12), fg=C["text_dim"],
                 bg=C["bg"]).pack(side="left", padx=10, pady=20)

        # dep badges top right
        self._dep_frame = tk.Frame(header, bg=C["bg"])
        self._dep_frame.pack(side="right", padx=20)

        # ── TILE ROW ────────────────────────────────────────────────────
        tile_bar = ctk.CTkFrame(self, fg_color="transparent", height=110)
        tile_bar.pack(fill="x", padx=18, pady=(0,4))
        tile_bar.pack_propagate(False)

        self._cat_tiles = {}
        for i, (cat, info) in enumerate(CONV_MAP.items()):
            tile = MetroTile(tile_bar, info["icon"], cat,
                             info["color"], size=(108, 90),
                             command=lambda c=cat: self._filter_by_cat(c))
            tile.grid(row=0, column=i, padx=3, pady=6, sticky="nsew")
            tile_bar.grid_columnconfigure(i, weight=1)
            self._cat_tiles[cat] = tile

        # ── MAIN BODY ───────────────────────────────────────────────────
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=18, pady=(0,12))
        body.grid_columnconfigure(0, weight=3)
        body.grid_columnconfigure(1, weight=0, minsize=230)
        body.grid_rowconfigure(0, weight=1)

        # LEFT – queue
        queue_outer = ctk.CTkFrame(body, corner_radius=0,
                                    fg_color=C["panel"],
                                    border_width=1,
                                    border_color=C["border"])
        queue_outer.grid(row=0, column=0, sticky="nsew", padx=(0,10))
        queue_outer.grid_rowconfigure(1, weight=1)
        queue_outer.grid_columnconfigure(0, weight=1)

        # queue header
        qh = ctk.CTkFrame(queue_outer, fg_color=C["border"],
                           corner_radius=0, height=34)
        qh.grid(row=0, column=0, sticky="ew")
        qh.pack_propagate(False)
        ctk.CTkLabel(qh, text="  FILE QUEUE", font=F(12,"bold"),
                     text_color=C["text_dim"],
                     fg_color="transparent").pack(side="left", padx=8)
        self._count_lbl = ctk.CTkLabel(qh, text="0 files", font=F(11),
                                        text_color=C["accent"],
                                        fg_color="transparent")
        self._count_lbl.pack(side="right", padx=12)

        # scrollable queue
        self._queue_scroll = ctk.CTkScrollableFrame(
            queue_outer, corner_radius=0, fg_color="transparent",
            scrollbar_fg_color=C["panel"],
            scrollbar_button_color=C["border"],
            scrollbar_button_hover_color=C["accent"])
        self._queue_scroll.grid(row=1, column=0, sticky="nsew")

        # empty state
        self._empty_lbl = tk.Label(
            self._queue_scroll,
            text="DROP FILES HERE\n\nor click Browse",
            font=(FONT_DISPLAY, 18),
            fg=C["text_darker"], bg=C["panel"],
            justify="center")
        self._empty_lbl.pack(expand=True, pady=80)

        # DROP target
        if DND_OK:
            self._queue_scroll.drop_target_register(DND_FILES)
            self._queue_scroll.dnd_bind("<<Drop>>", self._on_drop)

        # RIGHT – controls
        ctrl = ctk.CTkFrame(body, corner_radius=0,
                             fg_color=C["panel"],
                             border_width=1,
                             border_color=C["border"])
        ctrl.grid(row=0, column=1, sticky="nsew")

        self._build_controls(ctrl)

    def _build_controls(self, parent):
        pad = {"padx": 14, "pady": 6}

        # ── section: Add files ──────────────────────────────────────────
        self._sec(parent, "ADD FILES")
        ctk.CTkButton(parent, text="+ BROWSE FILES",
                      corner_radius=0, height=40,
                      font=F(13,"bold"),
                      fg_color=C["tile"],
                      hover_color=C["hover"],
                      command=self._browse).pack(fill="x", **pad)

        ctk.CTkButton(parent, text="CLEAR QUEUE",
                      corner_radius=0, height=34,
                      font=F(12),
                      fg_color="transparent",
                      border_width=1,
                      border_color=C["border"],
                      hover_color=C["hover"],
                      command=self._clear).pack(fill="x", **pad)

        self._sep(parent)

        # ── section: Format ─────────────────────────────────────────────
        self._sec(parent, "OUTPUT FORMAT")
        ctk.CTkLabel(parent, text="Convert to:", font=F(11),
                     text_color=C["text_dim"],
                     fg_color="transparent",
                     anchor="w").pack(fill="x", padx=14)

        self._fmt_var = ctk.StringVar(value="PDF")
        self._fmt_menu = ctk.CTkOptionMenu(
            parent,
            values=["PDF","JPG","PNG","MP4","MP3","WEBM","DOCX","XLSX",
                    "CSV","JSON","EPUB","ZIP","WAV","FLAC","WEBP","GIF"],
            variable=self._fmt_var,
            height=36, corner_radius=0,
            font=F(13,"bold"),
            fg_color=C["tile"],
            button_color=C["accent"],
            button_hover_color=C["selected"],
            dropdown_fg_color=C["panel"],
            dropdown_hover_color=C["hover"])
        self._fmt_menu.pack(fill="x", padx=14, pady=4)

        self._sep(parent)

        # ── section: Destination ────────────────────────────────────────
        self._sec(parent, "DESTINATION")
        ctk.CTkButton(parent, text="CHOOSE FOLDER",
                      corner_radius=0, height=38,
                      font=F(12,"bold"),
                      fg_color=C["accent"],
                      hover_color=C["selected"],
                      command=self._pick_output).pack(fill="x", **pad)

        self._dir_lbl = ctk.CTkLabel(parent, text="No folder selected",
                                      font=F(10),
                                      text_color=C["text_dim"],
                                      fg_color="transparent",
                                      wraplength=190, justify="left")
        self._dir_lbl.pack(fill="x", padx=14, pady=(0,4))

        # auto-open
        self._open_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(parent, text="Auto-open folder",
                        variable=self._open_var,
                        corner_radius=0, fg_color=C["accent"],
                        hover_color=C["selected"],
                        font=F(12)).pack(anchor="w", padx=14, pady=4)

        self._sep(parent)

        # ── Progress ────────────────────────────────────────────────────
        self._sec(parent, "PROGRESS")
        self._prog = MetroProgressBar(parent, color=C["accent"],
                                       height=8)
        self._prog.pack(fill="x", padx=14, pady=(4,2))
        self._prog_lbl = ctk.CTkLabel(parent, text="Ready",
                                       font=F(10), text_color=C["text_dim"],
                                       fg_color="transparent")
        self._prog_lbl.pack(anchor="w", padx=14)

        # spacer
        ctk.CTkLabel(parent, text="", fg_color="transparent"
                     ).pack(expand=True, fill="y")

        # ── BIG CONVERT BUTTON ──────────────────────────────────────────
        self._conv_btn = ctk.CTkButton(
            parent,
            text="START CONVERSION",
            corner_radius=0, height=56,
            font=F(14,"bold"),
            fg_color=C["tile"],
            hover_color=C["selected"],
            state="disabled",
            command=self._start_conversion)
        self._conv_btn.pack(fill="x", padx=14, pady=14, side="bottom")

    def _sec(self, parent, text):
        ctk.CTkLabel(parent, text=text,
                     font=F(10,"bold"),
                     text_color=C["accent"],
                     fg_color="transparent",
                     anchor="w").pack(fill="x", padx=14, pady=(12,0))

    def _sep(self, parent):
        ctk.CTkFrame(parent, height=1, fg_color=C["border"],
                     corner_radius=0).pack(fill="x", padx=10, pady=4)

    # ── Dependency badges ──────────────────────────────────────────────────

    def _check_deps(self):
        deps = [
            ("LO", _which("libreoffice") or os.path.exists(SOFFICE), "Docs"),
            ("FF", HAS_FF,     "Video/Audio"),
            ("PIL", PIL_OK,    "Images"),
            ("CAL", HAS_CAL,   "Ebooks"),
            ("HEIF", HEIF_OK,  "HEIC"),
        ]
        for tag, ok, tip in deps:
            bg  = C["accent4"] if ok else C["accent3"]
            lbl = tk.Label(self._dep_frame, text=f" {tag} ",
                           font=(FONT_MONO, 9, "bold"),
                           fg=C["text"], bg=bg, padx=4, pady=2)
            lbl.pack(side="left", padx=2)

    # ── File management ───────────────────────────────────────────────────

    def _browse(self):
        all_exts = " ".join(f"*{e}" for e in ALL_EXTS.keys())
        files = filedialog.askopenfilenames(
            title="Select files to convert",
            filetypes=[("All supported files", all_exts), ("All files","*.*")])
        if files:
            self._add_files(files)

    def _on_drop(self, event):
        raw   = event.data
        files = [p[0] or p[1]
                 for p in re.findall(r'\{([^}]+)\}|([^{}\s]+)', raw)]
        self._add_files(files)

    def _add_files(self, files):
        added = False
        for fp in files:
            if fp not in self._files:
                self._files.append(fp)
                row = QueueRow(self._queue_scroll, fp,
                               on_remove=self._remove_file)
                row.pack(fill="x", pady=1, padx=1)
                self._rows[fp] = row
                added = True
        if added:
            self._empty_lbl.pack_forget()
            self._update_format_suggestions()
        self._refresh_state()

    def _remove_file(self, fp):
        if fp in self._files:
            self._files.remove(fp)
        if fp in self._rows:
            self._rows[fp].destroy()
            del self._rows[fp]
        if not self._files:
            self._empty_lbl.pack(expand=True, pady=80)
        self._refresh_state()

    def _clear(self):
        for row in self._rows.values():
            row.destroy()
        self._rows.clear()
        self._files.clear()
        self._empty_lbl.pack(expand=True, pady=80)
        self._refresh_state()

    def _filter_by_cat(self, cat):
        """Clicking a tile filters the output format options to that category."""
        info = CONV_MAP[cat]
        self._fmt_menu.configure(values=info["outputs"])
        self._fmt_var.set(info["outputs"][0])
        # flash that tile
        tile = self._cat_tiles[cat]
        orig = tile._color
        tile.configure(fg_color=C["text"])
        self.after(120, lambda: tile.configure(fg_color=orig))

    def _update_format_suggestions(self):
        """Auto-suggest output formats based on queued files."""
        cats = set()
        for fp in self._files:
            ext = os.path.splitext(fp)[1].lower()
            c   = ALL_EXTS.get(ext)
            if c:
                cats.add(c)
        if len(cats) == 1:
            cat  = next(iter(cats))
            info = CONV_MAP[cat]
            self._fmt_menu.configure(values=info["outputs"])
            self._fmt_var.set(info["outputs"][0])
        elif len(cats) > 1:
            # mixed — show universal options
            self._fmt_menu.configure(
                values=["PDF","JPG","MP4","MP3","ZIP","JSON","CSV"])

    def _pick_output(self):
        folder = filedialog.askdirectory(title="Choose output folder")
        if folder:
            self._out_dir = folder
            display = folder if len(folder) <= 28 else f"…{folder[-25:]}"
            self._dir_lbl.configure(text=display, text_color=C["text"])
        self._refresh_state()

    def _refresh_state(self):
        n = len(self._files)
        self._count_lbl.configure(text=f"{n} file{'s' if n!=1 else ''}")
        ready = bool(self._files) and bool(self._out_dir) and not self._busy
        self._conv_btn.configure(
            state="normal" if ready else "disabled",
            fg_color=C["accent"] if ready else C["tile"])

    # ── Conversion ─────────────────────────────────────────────────────────

    def _start_conversion(self):
        if self._busy or not self._files or not self._out_dir:
            return
        self._busy = True
        self._conv_btn.configure(state="disabled", text="CONVERTING…",
                                  fg_color=C["accent5"])
        self._prog.set(0)
        self._prog.start_shimmer()
        self._prog_step = 1.0 / max(len(self._files), 1)
        self._prog_val  = 0.0
        self._success   = 0

        target_fmt = self._fmt_var.get()
        files_snap = list(self._files)
        threading.Thread(target=self._worker,
                         args=(files_snap, target_fmt),
                         daemon=True).start()

    def _worker(self, files, target_fmt):
        with concurrent.futures.ThreadPoolExecutor(
                max_workers=max(os.cpu_count() or 4, 4)) as ex:
            futs = {}
            for fp in files:
                ext = os.path.splitext(fp)[1].lower()
                cat = ALL_EXTS.get(ext)
                if cat:
                    engine = CONV_MAP[cat]["engine"]
                    self.after(0, lambda f=fp: self._rows[f].set_state("converting")
                               if f in self._rows else None)
                    fut = ex.submit(
                        ConversionEngine.convert,
                        fp, self._out_dir, target_fmt, engine)
                    futs[fut] = fp
                else:
                    self.after(0, lambda f=fp: self._rows[f].set_state("error")
                               if f in self._rows else None)

            for fut in concurrent.futures.as_completed(futs):
                fp  = futs[fut]
                ok  = fut.result()
                self._success += int(ok)
                state = "done" if ok else "error"
                self.after(0, lambda f=fp, s=state:
                           self._rows[f].set_state(s) if f in self._rows else None)
                self.after(0, self._tick_progress)

        self.after(0, self._finish, len(files))

    def _tick_progress(self):
        self._prog_val = min(1.0, self._prog_val + self._prog_step)
        self._prog.set(self._prog_val)
        n   = int(self._prog_val / self._prog_step)
        tot = int(1.0 / self._prog_step)
        self._prog_lbl.configure(text=f"{n} / {tot} done")

    def _finish(self, total):
        self._prog.set(1.0)
        self._prog.stop_shimmer()
        self._busy = False
        self._prog_lbl.configure(
            text=f"✔ {self._success}/{total} converted",
            text_color=C["accent4"] if self._success == total else C["accent3"])
        self._conv_btn.configure(
            text="START CONVERSION",
            fg_color=C["accent"] if (self._files and self._out_dir) else C["tile"],
            state="normal" if (self._files and self._out_dir) else "disabled")
        messagebox.showinfo(
            "OmniConvert",
            f"Converted {self._success} of {total} file(s) successfully.")
        if self._open_var.get() and self._out_dir:
            if sys.platform == "win32":
                os.startfile(self._out_dir)
            elif sys.platform == "darwin":
                subprocess.run(["open", self._out_dir])
            else:
                subprocess.run(["xdg-open", self._out_dir])


# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = OmniConvertApp()
    app.mainloop()