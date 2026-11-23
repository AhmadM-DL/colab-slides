"""Render PowerPoint / Google Slides decks inside a Colab (or Jupyter) notebook.

Two rendering engines:
  * "embed" - zero-install. Embeds the deck through an online viewer
               (Google Slides embed or Microsoft Office Online). Needs a
               publicly reachable link.
  * "local" - downloads the .pptx, converts it with LibreOffice -> PDF -> PNG
               inside the Colab VM and shows an interactive slideshow.
               Works for private Drive files (after drive.mount) and local paths.
"""
from __future__ import annotations

import base64
import html
import json
import os
import re
import shutil
import subprocess
import tempfile
import uuid
from pathlib import Path
from urllib.parse import parse_qs, quote, urlencode, urlparse, urlunparse

try:
    from IPython.display import HTML, IFrame, display
except ImportError:  # pragma: no cover
    HTML = IFrame = display = None

__all__ = ["show", "normalize_link", "render_local", "slide_images"]

_GSLIDES_RE = re.compile(r"docs\.google\.com/presentation/d/(?:e/)?([\w-]+)")
_GDRIVE_RE = re.compile(r"drive\.google\.com/(?:file/d/|open\?id=|uc\?.*id=)([\w-]+)")


# --------------------------------------------------------------------------- #
# Link handling
# --------------------------------------------------------------------------- #
def normalize_link(url: str) -> dict:
    """Classify a link and return the URLs each engine needs.

    Returns dict with keys: kind, embed_url, download_url.
    """
    url = url.strip()
    if not re.match(r"^https?://", url):
        return {"kind": "path", "embed_url": None, "download_url": url}

    m = _GSLIDES_RE.search(url)
    if m:
        fid = m.group(1)
        if "/presentation/d/e/" in url:  # "Publish to web" link
            base = f"https://docs.google.com/presentation/d/e/{fid}"
            return {"kind": "google_slides_published",
                    "embed_url": f"{base}/embed?start=false&loop=false",
                    "download_url": f"{base}/pub?output=pptx"}
        return {"kind": "google_slides",
                "embed_url": f"https://docs.google.com/presentation/d/{fid}/embed?start=false&loop=false",
                "download_url": f"https://docs.google.com/presentation/d/{fid}/export/pptx"}

    m = _GDRIVE_RE.search(url)
    if m:
        fid = m.group(1)
        return {"kind": "google_drive",
                "embed_url": f"https://drive.google.com/file/d/{fid}/preview",
                "download_url": f"https://drive.google.com/uc?export=download&id={fid}"}

    p = urlparse(url)
    if "dropbox.com" in p.netloc:
        q = parse_qs(p.query)
        q["dl"] = ["1"]
        dl = urlunparse(p._replace(query=urlencode(q, doseq=True)))
        q["raw"] = ["1"]
        q.pop("dl")
        raw = urlunparse(p._replace(query=urlencode(q, doseq=True)))
        return {"kind": "dropbox", "embed_url": _office_embed(raw), "download_url": dl}

    if "1drv.ms" in p.netloc or "onedrive.live.com" in p.netloc or "sharepoint.com" in p.netloc:
        # OneDrive/SharePoint share links already open in Office Online
        embed = url if "embed" in url else url + ("&" if "?" in url else "?") + "action=embedview"
        return {"kind": "onedrive", "embed_url": embed, "download_url": None}

    # Any other direct link to a .pptx/.ppt/.odp file
    return {"kind": "direct", "embed_url": _office_embed(url), "download_url": url}


def _office_embed(file_url: str) -> str:
    return "https://view.officeapps.live.com/op/embed.aspx?src=" + quote(file_url, safe="")


# --------------------------------------------------------------------------- #
# Local rendering (LibreOffice)
# --------------------------------------------------------------------------- #
def _ensure_tools(auto_install: bool = True) -> None:
    missing = [t for t in ("soffice", "pdftoppm") if not shutil.which(t)]
    if not missing:
        return
    if not auto_install:
        raise RuntimeError(f"Missing tools: {missing}. Install libreoffice-impress and poppler-utils.")
    print("Installing LibreOffice + poppler (one-time, ~1-2 min on Colab)...")
    subprocess.run("apt-get -qq update && apt-get -qq install -y libreoffice-impress poppler-utils "
                   "fonts-liberation fonts-dejavu > /dev/null", shell=True, check=True)


def _download(src: str, dest_dir: Path) -> Path:
    if not re.match(r"^https?://", src):
        path = Path(src).expanduser()
        if not path.exists():
            raise FileNotFoundError(path)
        return path
    import requests  # available on Colab

    sess = requests.Session()
    r = sess.get(src, stream=True, timeout=60, allow_redirects=True)
    # Google Drive large-file "virus scan" confirmation page
    if "text/html" in r.headers.get("content-type", "") and "drive.google.com" in src:
        m = re.search(r'name="confirm" value="([^"]+)"', r.text) or re.search(r"confirm=([\w-]+)", r.text)
        uuid_m = re.search(r'name="uuid" value="([^"]+)"', r.text)
        if m:
            params = {"export": "download", "confirm": m.group(1),
                      "id": parse_qs(urlparse(src).query)["id"][0]}
            if uuid_m:
                params["uuid"] = uuid_m.group(1)
            r = sess.get("https://drive.usercontent.google.com/download", params=params,
                         stream=True, timeout=60)
    r.raise_for_status()
    if "text/html" in r.headers.get("content-type", ""):
        raise RuntimeError("The link returned a web page, not a file. Make sure it is shared "
                           "as 'Anyone with the link', or mount Drive and pass the file path.")
    name = "deck"
    cd = r.headers.get("content-disposition", "")
    m = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^";]+)', cd)
    ext = Path(m.group(1)).suffix if m else Path(urlparse(src).path).suffix
    path = dest_dir / (name + (ext if ext.lower() in {".ppt", ".pptx", ".odp", ".pdf", ".key"} else ".pptx"))
    with open(path, "wb") as f:
        for chunk in r.iter_content(1 << 16):
            f.write(chunk)
    return path


def slide_images(src: str, dpi: int = 110, auto_install: bool = True) -> list[bytes]:
    """Return a list of PNG bytes, one per slide."""
    _ensure_tools(auto_install)
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        deck = _download(src, tmp)
        if deck.suffix.lower() == ".pdf":
            pdf = deck
        else:
            subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(tmp), str(deck)],
                           check=True, capture_output=True, timeout=300,
                           env={**os.environ, "HOME": str(tmp)})
            pdf = tmp / (deck.stem + ".pdf")
            if not pdf.exists():
                raise RuntimeError("LibreOffice conversion failed.")
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), str(pdf), str(tmp / "s")],
                       check=True, capture_output=True, timeout=300)
        files = sorted(tmp.glob("s-*.png"), key=lambda p: int(p.stem.split("-")[-1]))
        return [p.read_bytes() for p in files]


_PLAYER = """
<div id="{id}" class="cs-wrap" tabindex="0">
<style>
#{id}{{font-family:system-ui,sans-serif;max-width:{width}px;outline:none;user-select:none}}
#{id} .cs-stage{{position:relative;background:#111;border-radius:8px;overflow:hidden}}
#{id} .cs-stage img{{width:100%;display:block}}
#{id}:fullscreen{{max-width:none;background:#000;display:flex;flex-direction:column;justify-content:center}}
#{id}:fullscreen .cs-stage img{{max-height:92vh;object-fit:contain}}
#{id} .cs-bar{{display:flex;gap:8px;align-items:center;justify-content:center;padding:8px 0;color:#888}}
#{id} button{{border:1px solid #888;background:transparent;color:inherit;border-radius:6px;padding:4px 12px;cursor:pointer;font-size:14px}}
#{id} button:hover{{background:rgba(127,127,127,.2)}}
#{id} .cs-thumbs{{display:flex;gap:6px;overflow-x:auto;padding-bottom:6px}}
#{id} .cs-thumbs img{{height:54px;border:2px solid transparent;border-radius:4px;cursor:pointer;opacity:.6}}
#{id} .cs-thumbs img.on{{border-color:#4a8cff;opacity:1}}
#{id}:fullscreen .cs-thumbs{{display:none}}
</style>
<div class="cs-stage"><img class="cs-main" alt="slide"></div>
<div class="cs-bar">
  <button data-a="first">&#x23EE;</button><button data-a="prev">&#x25C0;</button>
  <span class="cs-count"></span>
  <button data-a="next">&#x25B6;</button><button data-a="last">&#x23ED;</button>
  <button data-a="full">&#x26F6; Fullscreen</button>
</div>
<div class="cs-thumbs"></div>
<script>
(function(){{
  const root=document.getElementById("{id}"), S={slides}; let i=0;
  const main=root.querySelector(".cs-main"), cnt=root.querySelector(".cs-count"), th=root.querySelector(".cs-thumbs");
  S.forEach((s,k)=>{{const t=document.createElement("img");t.src=s;t.onclick=()=>go(k);th.appendChild(t);}});
  function go(k){{i=Math.max(0,Math.min(S.length-1,k));main.src=S[i];cnt.textContent=(i+1)+" / "+S.length;
    [...th.children].forEach((t,n)=>t.classList.toggle("on",n===i));th.children[i].scrollIntoView({{block:"nearest",inline:"center"}});}}
  const act={{first:()=>go(0),prev:()=>go(i-1),next:()=>go(i+1),last:()=>go(S.length-1),
    full:()=>document.fullscreenElement?document.exitFullscreen():root.requestFullscreen()}};
  root.querySelectorAll("button").forEach(b=>b.onclick=()=>{{act[b.dataset.a]();root.focus();}});
  main.onclick=e=>{{const r=main.getBoundingClientRect();go(e.clientX-r.left<r.width/3?i-1:i+1);}};
  root.addEventListener("keydown",e=>{{const m={{ArrowRight:"next",ArrowDown:"next"," ":"next",PageDown:"next",
    ArrowLeft:"prev",ArrowUp:"prev",PageUp:"prev",Home:"first",End:"last",f:"full"}};
    if(m[e.key]){{e.preventDefault();act[m[e.key]]();}}}});
  go(0);
}})();
</script>
</div>
"""


def render_local(src: str, width: int = 900, dpi: int = 110, auto_install: bool = True) -> "HTML":
    imgs = slide_images(src, dpi=dpi, auto_install=auto_install)
    if not imgs:
        raise RuntimeError("No slides found.")
    uris = ["data:image/png;base64," + base64.b64encode(b).decode() for b in imgs]
    return HTML(_PLAYER.format(id="cs" + uuid.uuid4().hex[:8], width=width, slides=json.dumps(uris)))


def _render_embed(embed_url: str, width: int, height: int) -> "HTML":
    u = html.escape(embed_url, quote=True)
    return HTML(f'<iframe src="{u}" width="{width}" height="{height}" frameborder="0" '
                f'allowfullscreen="true" mozallowfullscreen="true" webkitallowfullscreen="true" '
                f'style="border-radius:8px;max-width:100%"></iframe>'
                f'<div style="font:12px sans-serif;color:#888">Not loading? Try '
                f'<code>show(url, mode="local")</code></div>')


# --------------------------------------------------------------------------- #
# Public entry point
# --------------------------------------------------------------------------- #
def show(url: str, mode: str = "auto", width: int = 900, height: int | None = None,
         dpi: int = 110, auto_install: bool = True):
    """Render a slide deck in the notebook output.

    url   : Google Slides / Drive / Dropbox / OneDrive / direct .pptx link, or a local path.
    mode  : "auto" (embed when possible, else local), "embed", or "local".
    """
    info = normalize_link(url)
    height = height or int(width * 9 / 16) + 30
    if mode == "auto":
        mode = "embed" if info["embed_url"] else "local"
    if mode == "embed":
        if not info["embed_url"]:
            raise ValueError("This link can't be embedded; use mode='local'.")
        out = _render_embed(info["embed_url"], width, height)
    elif mode == "local":
        src = info["download_url"]
        if not src:
            raise ValueError("No downloadable URL for this link; share it as a direct download link.")
        out = render_local(src, width=width, dpi=dpi, auto_install=auto_install)
    else:
        raise ValueError("mode must be 'auto', 'embed' or 'local'")
    display(out)
