import shutil

import pytest

import colab_slides
from colab_slides.core import normalize_link


@pytest.mark.parametrize("url,kind,embed_part,download_part", [
    ("https://docs.google.com/presentation/d/ABC123/edit#slide=id.p", "google_slides",
     "/d/ABC123/embed", "/d/ABC123/export/pptx"),
    ("https://docs.google.com/presentation/d/e/2PACX-xyz/pub?start=false", "google_slides_published",
     "/d/e/2PACX-xyz/embed", "pub?output=pptx"),
    ("https://drive.google.com/file/d/FID9/view?usp=sharing", "google_drive",
     "/file/d/FID9/preview", "id=FID9"),
    ("https://www.dropbox.com/s/abc/deck.pptx?dl=0", "dropbox", "view.officeapps.live.com", "dl=1"),
    ("https://example.com/talk.pptx", "direct", "view.officeapps.live.com", "example.com/talk.pptx"),
])
def test_normalize_link(url, kind, embed_part, download_part):
    info = normalize_link(url)
    assert info["kind"] == kind
    assert embed_part in info["embed_url"]
    assert download_part in info["download_url"]


def test_local_path_is_not_embeddable():
    info = normalize_link("/content/deck.pptx")
    assert info == {"kind": "path", "embed_url": None, "download_url": "/content/deck.pptx"}


@pytest.mark.skipif(not (shutil.which("soffice") and shutil.which("pdftoppm")),
                    reason="LibreOffice/poppler not installed")
def test_local_render(tmp_path):
    from pptx import Presentation
    prs = Presentation()
    for i in range(3):
        prs.slides.add_slide(prs.slide_layouts[1]).shapes.title.text = f"Slide {i + 1}"
    deck = tmp_path / "deck.pptx"
    prs.save(deck)

    imgs = colab_slides.slide_images(str(deck), auto_install=False)
    assert len(imgs) == 3
    assert all(b.startswith(b"\x89PNG") for b in imgs)

    html = colab_slides.render_local(str(deck), auto_install=False).data
    assert html.count("data:image/png;base64,") == 3
