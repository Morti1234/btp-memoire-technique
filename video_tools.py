"""
Outils vidéo MarketBuilding
- Optimisation web des vidéos lourdes (.mov Retina -> .mp4 H.264 léger, lecture instantanée)
- Génération automatique de vidéos de présentation (intro + démo + outro, fondus, bandeau)
Dépendances : FFmpeg (brew install ffmpeg) + Pillow
"""

import json
import re
import shutil
import subprocess
import tempfile
import time
from pathlib import Path
from typing import Callable, Optional

from PIL import Image, ImageDraw, ImageFont

BASE_DIR = Path(__file__).resolve().parent
VIDEOS_DIR = BASE_DIR / "videos"
META_FILE = VIDEOS_DIR / "meta.json"
FONT_DIR = BASE_DIR / "fonts"

FORMATS = {
    "16:9 - Site web / YouTube (1920x1080)": (1920, 1080),
    "9:16 - Reels / TikTok / Shorts (1080x1920)": (1080, 1920),
    "1:1 - LinkedIn / Instagram (1080x1080)": (1080, 1080),
}

BG = (14, 17, 23)
NAVY = (30, 58, 138)
BLUE = (59, 130, 246)
ORANGE = (245, 158, 11)
WHITE = (255, 255, 255)
GREY = (148, 163, 184)
FPS = 30
FADE = 0.6

ProgressFn = Optional[Callable[[float, str], None]]


# ─────────────────────────────────────────────────────────────────────────
# Utilitaires
# ─────────────────────────────────────────────────────────────────────────
def ffmpeg_bin() -> str:
    path = shutil.which("ffmpeg") or "/opt/homebrew/bin/ffmpeg"
    if not Path(path).exists():
        raise RuntimeError("FFmpeg introuvable. Installez-le avec : brew install ffmpeg")
    return path


def media_info(path: Path) -> dict:
    """Durée, dimensions et présence d'audio (via la sortie de ffmpeg -i)."""
    res = subprocess.run([ffmpeg_bin(), "-hide_banner", "-i", str(path)], capture_output=True, text=True)
    err = res.stderr
    info = {"duration": 0.0, "width": 0, "height": 0, "has_audio": "Audio:" in err}
    m = re.search(r"Duration:\s*(\d+):(\d+):(\d+(?:\.\d+)?)", err)
    if m:
        h, mi, s = m.groups()
        info["duration"] = int(h) * 3600 + int(mi) * 60 + float(s)
    m = re.search(r"Video:.*?(\d{2,5})x(\d{2,5})", err)
    if m:
        info["width"], info["height"] = int(m.group(1)), int(m.group(2))
    return info


def _run_ffmpeg(cmd: list, total_seconds: float, progress: ProgressFn, label: str) -> None:
    """Lance FFmpeg en suivant la progression (-progress pipe:1)."""
    cmd = [cmd[0], "-hide_banner", "-nostats", "-progress", "pipe:1"] + cmd[1:]
    with tempfile.TemporaryFile(mode="w+") as errf:
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=errf, text=True)
        for line in proc.stdout:
            if line.startswith("out_time_us=") or line.startswith("out_time_ms="):
                try:
                    secs = int(line.split("=")[1]) / 1_000_000
                except ValueError:
                    continue
                if progress and total_seconds > 0:
                    progress(min(secs / total_seconds, 0.99), label)
        proc.wait()
        if proc.returncode != 0:
            errf.seek(0)
            raise RuntimeError("Erreur FFmpeg :\n" + errf.read()[-2000:])
    if progress:
        progress(1.0, label)


def _safe_name(name: str) -> str:
    stem = re.sub(r"[^\w\-]+", "_", Path(name).stem, flags=re.UNICODE).strip("_")
    return (stem or "video")[:60]


# ─────────────────────────────────────────────────────────────────────────
# Métadonnées (titres, vidéo mise en avant)
# ─────────────────────────────────────────────────────────────────────────
def load_meta() -> dict:
    VIDEOS_DIR.mkdir(exist_ok=True)
    if META_FILE.exists():
        try:
            return json.loads(META_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"featured": None, "titles": {}}


def save_meta(meta: dict) -> None:
    VIDEOS_DIR.mkdir(exist_ok=True)
    META_FILE.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")


def list_videos() -> list:
    VIDEOS_DIR.mkdir(exist_ok=True)
    files = [p for p in VIDEOS_DIR.glob("*.mp4") if not p.name.startswith(".")]
    return sorted(files, key=lambda p: p.stat().st_mtime, reverse=True)


def featured_video() -> Optional[Path]:
    meta = load_meta()
    if meta.get("featured"):
        p = VIDEOS_DIR / meta["featured"]
        if p.exists():
            return p
    return None


def delete_video(path: Path) -> None:
    meta = load_meta()
    if meta.get("featured") == path.name:
        meta["featured"] = None
    meta.get("titles", {}).pop(path.name, None)
    save_meta(meta)
    path.unlink(missing_ok=True)


# ─────────────────────────────────────────────────────────────────────────
# 1. Optimisation web
# ─────────────────────────────────────────────────────────────────────────
def optimize_for_web(src: Path, title: str = "", progress: ProgressFn = None) -> Path:
    """Convertit n'importe quelle vidéo (.mov, .mp4, .webm…) en MP4 H.264 léger, max 1920 px."""
    VIDEOS_DIR.mkdir(exist_ok=True)
    info = media_info(src)
    dst = VIDEOS_DIR / f"{_safe_name(title or src.name)}_{time.strftime('%Y%m%d_%H%M%S')}.mp4"
    cmd = [
        ffmpeg_bin(), "-y", "-i", str(src),
        "-map", "0:v:0", "-map", "0:a:0?",
        "-vf", f"scale='min(1920,iw)':-2:flags=lanczos,fps={FPS},format=yuv420p",
        "-c:v", "libx264", "-preset", "veryfast", "-crf", "24",
        "-c:a", "aac", "-b:a", "128k",
        "-movflags", "+faststart", str(dst),
    ]
    _run_ffmpeg(cmd, info["duration"], progress, "Optimisation pour le web")
    if title:
        meta = load_meta()
        meta.setdefault("titles", {})[dst.name] = title
        save_meta(meta)
    return dst


# ─────────────────────────────────────────────────────────────────────────
# 2. Cartes visuelles (Pillow)
# ─────────────────────────────────────────────────────────────────────────
def _font(size: int, bold: bool = False):
    name = "CustomFont-Bold.ttf" if bold else "CustomFont-Regular.ttf"
    try:
        return ImageFont.truetype(str(FONT_DIR / name), size)
    except Exception:
        return ImageFont.load_default()


def _wrap(draw, text: str, font, max_w: int) -> list:
    lines = []
    for para in text.split("\n"):
        words, cur = para.split(), ""
        for w in words:
            test = f"{cur} {w}".strip()
            if draw.textlength(test, font=font) <= max_w:
                cur = test
            else:
                if cur:
                    lines.append(cur)
                cur = w
        lines.append(cur)
    return [l for l in lines if l is not None]


def _draw_brand(draw, x: int, y: int, scale: float) -> None:
    """Petit logo : skyline + grue + 'MarketBuilding'."""
    s = scale
    draw.rectangle([x, y + 30 * s, x + 18 * s, y + 70 * s], fill=BLUE)
    draw.rectangle([x + 22 * s, y + 10 * s, x + 46 * s, y + 70 * s], fill=(37, 99, 235))
    draw.rectangle([x + 50 * s, y + 24 * s, x + 70 * s, y + 70 * s], fill=(96, 165, 250))
    draw.line([x + 76 * s, y + 70 * s, x + 76 * s, y], fill=ORANGE, width=max(2, int(3 * s)))
    draw.line([x + 52 * s, y, x + 88 * s, y], fill=ORANGE, width=max(2, int(3 * s)))
    draw.text((x + 100 * s, y + 8 * s), "Market", font=_font(int(30 * s), True), fill=WHITE)
    draw.text((x + 100 * s, y + 40 * s), "Building", font=_font(int(30 * s)), fill=GREY)


def make_card(path: Path, size: tuple, title: str, subtitle: str = "", footer: str = "") -> Path:
    W, H = size
    img = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)
    # dégradé vertical discret
    for yy in range(H):
        t = yy / H
        c = tuple(int(BG[i] * (1 - t * 0.6) + NAVY[i] * t * 0.6) for i in range(3))
        draw.line([(0, yy), (W, yy)], fill=c)

    unit = min(W, H) / 1080
    margin = int(90 * unit)
    _draw_brand(draw, margin, margin, 1.3 * unit)

    max_w = W - 2 * margin
    f_title = _font(int(82 * unit), True)
    f_sub = _font(int(42 * unit))
    t_lines = _wrap(draw, title, f_title, max_w)
    s_lines = _wrap(draw, subtitle, f_sub, max_w) if subtitle else []
    lh_t, lh_s = int(100 * unit), int(58 * unit)
    block_h = len(t_lines) * lh_t + (int(64 * unit) + len(s_lines) * lh_s if s_lines else 0)
    y = (H - block_h) // 2

    for line in t_lines:
        draw.text(((W - draw.textlength(line, font=f_title)) / 2, y), line, font=f_title, fill=WHITE)
        y += lh_t
    bar_w = int(160 * unit)
    draw.rectangle([(W - bar_w) / 2, y + 24 * unit, (W + bar_w) / 2, y + 30 * unit], fill=ORANGE)
    y += int(64 * unit)
    for line in s_lines:
        draw.text(((W - draw.textlength(line, font=f_sub)) / 2, y), line, font=f_sub, fill=GREY)
        y += lh_s

    if footer:
        f_foot = _font(int(34 * unit), True)
        for i, line in enumerate(reversed(_wrap(draw, footer, f_foot, max_w))):
            yy = H - margin - (i + 1) * int(48 * unit)
            draw.text(((W - draw.textlength(line, font=f_foot)) / 2, yy), line, font=f_foot, fill=BLUE)

    img.save(path)
    return path


def make_caption(path: Path, size: tuple, text: str) -> Path:
    """Bandeau transparent en bas de la démo."""
    W, H = size
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    unit = min(W, H) / 1080
    f = _font(int(40 * unit), True)
    lines = _wrap(draw, text, f, W - int(160 * unit))
    lh = int(54 * unit)
    box_h = len(lines) * lh + int(40 * unit)
    y0 = H - box_h - int(60 * unit)
    draw.rounded_rectangle([int(40 * unit), y0, W - int(40 * unit), y0 + box_h], radius=int(18 * unit), fill=(14, 17, 23, 215))
    draw.rectangle([int(40 * unit), y0, int(52 * unit), y0 + box_h], fill=ORANGE)
    y = y0 + int(20 * unit)
    for line in lines:
        draw.text((int(80 * unit), y), line, font=f, fill=WHITE)
        y += lh
    img.save(path)
    return path


# ─────────────────────────────────────────────────────────────────────────
# 3. Génération de la vidéo de présentation
# ─────────────────────────────────────────────────────────────────────────
def generate_presentation(
    src: Path,
    fmt_label: str,
    title: str,
    subtitle: str,
    caption: str,
    outro_title: str,
    outro_contact: str,
    start: float = 0.0,
    end: Optional[float] = None,
    speed: float = 1.0,
    intro_s: float = 3.5,
    outro_s: float = 4.0,
    audio: Optional[Path] = None,
    progress: ProgressFn = None,
) -> Path:
    W, H = FORMATS[fmt_label]
    info = media_info(src)
    end = min(end or info["duration"], info["duration"])
    if end - start < 1:
        raise ValueError("L'extrait sélectionné est trop court (minimum 1 seconde).")
    clip_d = (end - start) / speed
    total = intro_s + clip_d + outro_s - 2 * FADE

    VIDEOS_DIR.mkdir(exist_ok=True)
    out = VIDEOS_DIR / f"Presentation_{_safe_name(title)}_{W}x{H}_{time.strftime('%Y%m%d_%H%M%S')}.mp4"

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        if progress:
            progress(0.02, "Création des visuels")
        intro = make_card(tmp / "intro.png", (W, H), title, subtitle)
        outro = make_card(tmp / "outro.png", (W, H), outro_title, outro_contact, footer="MarketBuilding")

        cmd = [
            ffmpeg_bin(), "-y",
            "-loop", "1", "-framerate", str(FPS), "-t", str(intro_s), "-i", str(intro),
            "-ss", str(start), "-to", str(end), "-i", str(src),
            "-loop", "1", "-framerate", str(FPS), "-t", str(outro_s), "-i", str(outro),
        ]
        card = f"scale={W}:{H},setsar=1,fps={FPS},format=yuv420p"
        demo = (
            f"[1:v]setpts=(PTS-STARTPTS)/{speed},"
            f"scale={W}:{H}:force_original_aspect_ratio=decrease:flags=lanczos,"
            f"pad={W}:{H}:(ow-iw)/2:(oh-ih)/2:color=0x0e1117,setsar=1,fps={FPS},format=yuv420p"
        )
        filters = [f"[0:v]{card}[v0]", f"[2:v]{card}[v2]"]
        idx = 3
        if caption.strip():
            cap = make_caption(tmp / "caption.png", (W, H), caption.strip())
            cmd += ["-loop", "1", "-i", str(cap)]
            filters.append(f"{demo}[d0]")
            filters.append(f"[d0][{idx}:v]overlay=0:0:shortest=1,format=yuv420p[v1]")
            idx += 1
        else:
            filters.append(f"{demo}[v1]")
        filters.append(f"[v0][v1]xfade=transition=fade:duration={FADE}:offset={intro_s - FADE:.3f}[x1]")
        filters.append(f"[x1][v2]xfade=transition=fade:duration={FADE}:offset={intro_s + clip_d - 2 * FADE:.3f}[vout]")

        maps = ["-map", "[vout]"]
        if audio:
            cmd += ["-i", str(audio)]
            filters.append(f"[{idx}:a]atrim=0:{total:.3f},afade=t=out:st={max(total - 2, 0):.3f}:d=2[aout]")
            maps += ["-map", "[aout]", "-c:a", "aac", "-b:a", "160k"]

        cmd += ["-filter_complex", ";".join(filters)] + maps + [
            "-t", f"{total:.3f}",
            "-c:v", "libx264", "-preset", "veryfast", "-crf", "22",
            "-pix_fmt", "yuv420p", "-movflags", "+faststart", str(out),
        ]
        _run_ffmpeg(cmd, total, progress, "Montage de la vidéo")

    meta = load_meta()
    meta.setdefault("titles", {})[out.name] = f"{title} ({W}x{H})"
    save_meta(meta)
    return out
