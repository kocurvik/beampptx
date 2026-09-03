"""beampptx — convert LaTeX Beamer slides to PowerPoint.

Slides are embedded as full-bleed vector graphics, Beamer overlays become
individual slides, internal navigation becomes PowerPoint click actions, and
videos included with ``movie15`` or ``multimedia`` become native movie shapes.
"""

from importlib.metadata import PackageNotFoundError, version as _version

try:
    __version__ = _version("beampptx")
except PackageNotFoundError:  # running from a source checkout without install
    __version__ = "0.0.0.dev0"

__all__ = ["__version__", "main"]


def main() -> None:
    """Entry point for the ``beampptx`` command."""
    from .cli import main as _main

    _main()
