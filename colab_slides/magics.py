import argparse
import shlex

from IPython.core.magic import Magics, line_magic, magics_class

from .core import show


@magics_class
class SlidesMagics(Magics):
    @line_magic
    def slides(self, line):
        """%slides URL_OR_PATH [--local | --embed] [--width 900] [--height H] [--dpi 110]"""
        p = argparse.ArgumentParser(prog="%slides", add_help=False)
        p.add_argument("url")
        g = p.add_mutually_exclusive_group()
        g.add_argument("--local", action="store_true")
        g.add_argument("--embed", action="store_true")
        p.add_argument("--width", type=int, default=900)
        p.add_argument("--height", type=int, default=None)
        p.add_argument("--dpi", type=int, default=110)
        try:
            a = p.parse_args(shlex.split(self.shell.var_expand(line)))
        except SystemExit:
            print(self.slides.__doc__)
            return
        mode = "local" if a.local else "embed" if a.embed else "auto"
        show(a.url, mode=mode, width=a.width, height=a.height, dpi=a.dpi)
