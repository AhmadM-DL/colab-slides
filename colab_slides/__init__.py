"""colab_slides - show PowerPoint / Google Slides decks inside Colab.

    import colab_slides
    colab_slides.show("https://docs.google.com/presentation/d/<id>/edit")

or, with the magic:

    %load_ext colab_slides
    %slides https://example.com/deck.pptx --local
"""
from .core import normalize_link, render_local, show, slide_images

__version__ = "0.1.1"
__all__ = ["show", "normalize_link", "render_local", "slide_images", "load_ipython_extension"]


def load_ipython_extension(ipython):
    from .magics import SlidesMagics
    ipython.register_magics(SlidesMagics)
