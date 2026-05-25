#!/usr/bin/env python3
"""
beamer_to_pptx.py

Compiles a Beamer LaTeX file to PDF and converts each slide to a .pptx file,
embedding every slide as a full-bleed EMF/SVG vector graphic.

Usage:
    python beamer_to_pptx.py presentation.tex
    python beamer_to_pptx.py presentation.tex --output out.pptx
    python beamer_to_pptx.py presentation.tex --latex-engine xelatex

Requirements:
    pip install python-pptx pymupdf
    System: pdflatex (or xelatex/lualatex).
            biber and/or bibtex if the document has a bibliography.
"""

import argparse
import io
import os
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

import fitz  # PyMuPDF
import lxml.etree as etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

try:
    from pptx.opc.part import Part as _Part  # pptx < 1.0
    from pptx.opc.packuri import PackURI

    def Part(partname, content_type, blob, package):
        # pptx < 1.0: Part(partname, content_type, blob, package=None)
        return _Part(partname, content_type, blob, package)
except ImportError:
    from pptx.opc.package import Part as _Part, PackURI  # pptx >= 1.0

    def Part(partname, content_type, blob, package):
        # pptx >= 1.0: Part(partname, content_type, package, blob=b"")
        return _Part(partname, content_type, package, blob)


# ──────────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────────

def run(cmd: list[str], cwd: Path | None = None, check: bool = True,
        env: dict | None = None) -> subprocess.CompletedProcess:
    """Run a subprocess command, streaming output to the terminal."""
    print(f"  $ {' '.join(str(c) for c in cmd)}")
    return subprocess.run(cmd, cwd=cwd, check=check, env=env,
                          stdout=sys.stdout, stderr=sys.stderr)


def require(tool: str) -> str:
    """Return full path of *tool* or abort with a helpful message."""
    path = shutil.which(tool)
    if not path:
        sys.exit(f"[error] '{tool}' not found on PATH. Please install it and retry.")
    return path


# ──────────────────────────────────────────────────────────────────────────────
# Step 1 – Compile .tex → .pdf
# ──────────────────────────────────────────────────────────────────────────────

def compile_latex(tex_path: Path, engine: str, build_dir: Path,
                  latex_cwd: Path) -> Path:
    """
    Run the LaTeX engine (plus biber/bibtex when a bibliography is present)
    with the working directory set to *latex_cwd* so relative ``\\input``,
    ``\\includegraphics``, and ``\\bibliography`` paths resolve correctly.
    Output artefacts are written to *build_dir*.  Returns the path to the PDF.
    """
    engine_bin = require(engine)
    tex_abs = tex_path.resolve()
    stem = tex_abs.stem

    compile_cmd = [
        engine_bin,
        "-interaction=nonstopmode",
        "-halt-on-error",
        f"-output-directory={build_dir}",
        str(tex_abs),
    ]

    print(f"  LaTeX working directory: {latex_cwd}")
    print("\n[1/3] Compiling LaTeX (pass 1) …")
    run(compile_cmd, cwd=latex_cwd)

    _process_bibliography(build_dir, stem, latex_cwd)

    print("\n[1/3] Compiling LaTeX (pass 2) …")
    run(compile_cmd, cwd=latex_cwd)
    print("\n[1/3] Compiling LaTeX (pass 3 – cross-refs) …")
    run(compile_cmd, cwd=latex_cwd)

    pdf_path = build_dir / (stem + ".pdf")
    if not pdf_path.exists():
        sys.exit(f"[error] Expected PDF not found: {pdf_path}")
    print(f"  → PDF: {pdf_path}")
    return pdf_path


def _process_bibliography(build_dir: Path, stem: str, latex_cwd: Path) -> None:
    """
    Run biber or bibtex if the first LaTeX pass produced bibliography artefacts.

      *.bcf present  → biblatex + biber backend
      \\bibdata in .aux → traditional bibtex / natbib
      neither        → no bibliography, nothing to do

    Both tools run with cwd=*latex_cwd* (so relative .bib paths resolve the
    same way they did for LaTeX) and receive an absolute path to the
    .aux/.bcf file inside *build_dir*.  ``BIBINPUTS``/``BSTINPUTS`` are
    extended to also search *build_dir* so that biblatex's auto-generated
    ``<stem>-blx.bib`` (written next to the .aux) is picked up.  The trailing
    empty path element preserves the default TeX search locations.
    """
    bcf_path = build_dir / f"{stem}.bcf"
    aux_path = build_dir / f"{stem}.aux"
    target = str(build_dir / stem)

    search = os.pathsep.join([str(build_dir), str(latex_cwd), ""])
    env = {**os.environ, "BIBINPUTS": search, "BSTINPUTS": search}

    if bcf_path.exists():
        biber_bin = require("biber")
        print("\n[1/3] Processing bibliography (biber) …")
        run([biber_bin, target], cwd=latex_cwd, env=env)
        return

    if aux_path.exists() and "\\bibdata" in aux_path.read_text(errors="ignore"):
        bibtex_bin = require("bibtex")
        print("\n[1/3] Processing bibliography (bibtex) …")
        run([bibtex_bin, target], cwd=latex_cwd, env=env)


# ──────────────────────────────────────────────────────────────────────────────
# Step 2 – Split PDF into per-slide SVG vector files
# ──────────────────────────────────────────────────────────────────────────────

def pdf_to_svgs_pymupdf(pdf_path: Path, svg_dir: Path) -> list[Path]:
    """
    Use PyMuPDF (fitz) to render each PDF page as an SVG.
    PyMuPDF produces true vector SVGs (text + paths), not bitmaps.
    """
    svg_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(pdf_path))
    paths = []
    for i, page in enumerate(doc):
        svg_bytes = page.get_svg_image(matrix=fitz.Identity)
        out = svg_dir / f"slide_{i+1:04d}.svg"
        out.write_bytes(svg_bytes if isinstance(svg_bytes, bytes)
                        else svg_bytes.encode())
        paths.append(out)
    doc.close()
    print(f"  → {len(paths)} SVG(s) written to {svg_dir}")
    return paths


def extract_svgs(pdf_path: Path, svg_dir: Path) -> list[Path]:
    """Convert PDF pages to SVG using PyMuPDF."""
    print("\n[2/3] Converting PDF pages to SVG …")

    paths = pdf_to_svgs_pymupdf(pdf_path, svg_dir)
    if paths:
        return paths

    sys.exit(
        "[error] PyMuPDF is required for SVG extraction.\n"
        "Install it with:\n"
        "  pip install pymupdf"
    )


# ──────────────────────────────────────────────────────────────────────────────
# Step 3 – Build .pptx
# ──────────────────────────────────────────────────────────────────────────────

def _svg_dimensions(svg_path: Path) -> tuple[float, float]:
    """
    Parse width/height from the <svg> root element (in pt).
    Falls back to 4:3 Beamer default (362.835 × 272.126 pt ≈ 128×96 mm).
    """
    BEAMER_W_PT = 362.835   # 128 mm
    BEAMER_H_PT = 272.126   # 96 mm

    try:
        tree = ET.parse(str(svg_path))
        root = tree.getroot()
        # Strip namespace if present
        attrib = root.attrib

        def to_pt(value: str) -> float:
            value = value.strip()
            if value.endswith("pt"):
                return float(value[:-2])
            if value.endswith("mm"):
                return float(value[:-2]) * 2.8346456693
            if value.endswith("cm"):
                return float(value[:-2]) * 28.346456693
            if value.endswith("in"):
                return float(value[:-2]) * 72.0
            if value.endswith("px"):
                return float(value[:-2]) * 0.75  # 96 dpi → pt
            return float(value)  # assume pt

        w = to_pt(attrib["width"])  if "width"  in attrib else BEAMER_W_PT
        h = to_pt(attrib["height"]) if "height" in attrib else BEAMER_H_PT
        return w, h

    except Exception:
        return BEAMER_W_PT, BEAMER_H_PT


def build_pptx(svg_paths: list[Path], output_path: Path) -> None:
    """Create a .pptx where each slide contains one full-bleed SVG picture."""
    print("\n[3/3] Building PPTX …")

    # Infer slide dimensions from the first SVG
    slide_w_pt, slide_h_pt = _svg_dimensions(svg_paths[0])
    print(f"  Slide dimensions: {slide_w_pt:.2f} × {slide_h_pt:.2f} pt")

    prs = Presentation()

    # Set slide size to match the source PDF pages (EMU = pt × 12700)
    prs.slide_width  = int(slide_w_pt * 12700)
    prs.slide_height = int(slide_h_pt * 12700)

    # Use a completely blank layout (no placeholders)
    blank_layout = prs.slide_layouts[6]

    for idx, svg_path in enumerate(svg_paths, start=1):
        slide = prs.slides.add_slide(blank_layout)

        # python-pptx does not natively support SVG insertion into the
        # slide XML, so we add it via direct XML manipulation (Office 2016+
        # supports <a:graphicData uri="…svg…"> inside a picture shape).
        # For maximum compatibility we also embed a PNG fallback.
        _insert_svg_picture(slide, svg_path,
                            prs.slide_width, prs.slide_height)

        if idx % 10 == 0 or idx == len(svg_paths):
            print(f"  Processed {idx}/{len(svg_paths)} slides …")

    prs.save(str(output_path))
    print(f"\n✓ Saved: {output_path}")


# ──────────────────────────────────────────────────────────────────────────────
# SVG embedding helpers
# ──────────────────────────────────────────────────────────────────────────────

def _png_fallback(svg_path: Path) -> bytes:
    """
    Render SVG → PNG bytes.

    IMPORTANT: <a:blip r:embed="..."> pointing at a raster image is *required*
    by the OOXML spec — the <asvg:svgBlip> is only an extension on top of it.
    PowerPoint silently ignores the SVG and falls back to rasterising (or shows
    nothing) when <a:blip> has no r:embed.  This function therefore always
    returns valid PNG bytes; it aborts the process if every renderer fails.

    Uses PyMuPDF (fitz) to rasterise the SVG.
    """
    try:
        doc = fitz.open(stream=svg_path.read_bytes(), filetype="svg")
        pix = doc[0].get_pixmap(matrix=fitz.Matrix(1, 1))
        png = pix.tobytes("png")
        doc.close()
        return png
    except Exception as exc:
        sys.exit(
            f"[error] Cannot render SVG → PNG via PyMuPDF: {exc}\n"
            "A raster fallback is required for PowerPoint to honour the "
            "embedded SVG."
        )


def _insert_svg_picture(slide, svg_path: Path, cx_emu: int, cy_emu: int) -> None:
    """
    Embed *svg_path* as a full-slide picture using the Office Open XML SVG
    extension (DrawingML / OOXML §20.1.8.55).

    OOXML structure required by PowerPoint to actually honour the SVG:

      <a:blip r:embed="PNG_rId">              ← MANDATORY raster base image
        <a:extLst>
          <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
            <asvg:svgBlip r:embed="SVG_rId"/> ← vector overlay (Office 2016+)
          </a:ext>
        </a:extLst>
      </a:blip>

    The ``{96DAC541-…}`` URI is the well-known DrawingML extension identifier
    for SVG-aware blips.  Without the <a:extLst>/<a:ext> wrapper PowerPoint
    silently drops the <asvg:svgBlip> and renders only the PNG fallback.
    """
    svg_bytes = svg_path.read_bytes()
    # Always produces bytes or aborts — see _png_fallback docstring.
    png_bytes = _png_fallback(svg_path)

    slide_part = slide.part
    package = slide_part.package

    # PNG part — registered as a normal image relationship (r:embed on <a:blip>)
    try:
        _png_part, png_rId = slide_part.get_or_add_image_part(io.BytesIO(png_bytes))
    except Exception:
        png_partname = PackURI(f"/ppt/media/{svg_path.stem}_{id(svg_path)}_fb.png")
        png_part = Part(png_partname, "image/png", png_bytes, package)
        png_rId = slide_part.relate_to(png_part, RT.IMAGE)

    # SVG part — must use the *standard* image relationship type, just like the
    # PNG.  PowerPoint-generated pptx files always relate SVG attachments via
    # RT.IMAGE; the bespoke ``…/relationships/svgBlip`` URI is not recognised
    # by the renderer, which then refuses to load the picture entirely.
    # The ``package=`` backref is required so the part is correctly serialised
    # into the .pptx archive (without it, PowerPoint's resolver fails and the
    # whole picture shape collapses into a “Picture can't be displayed” box).
    svg_partname = PackURI(f"/ppt/media/{svg_path.stem}_{id(svg_path)}.svg")
    svg_part = Part(svg_partname, "image/svg+xml", svg_bytes, package)
    svg_rId = slide_part.relate_to(svg_part, RT.IMAGE)

    # ── Build the whole <p:pic> as one piece of XML so lxml uses the prefixes
    # declared on the root element (otherwise it invents ns0/ns1/… prefixes on
    # the <a:blip> subtree, which some readers refuse to load).
    SVG_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
    pic_xml = f"""\
<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
  <p:nvPicPr>
    <p:cNvPr id="2" name="Slide {svg_path.stem}"/>
    <p:cNvPicPr preferRelativeResize="0"/>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="{png_rId}">
      <a:extLst>
        <a:ext uri="{SVG_EXT_URI}">
          <asvg:svgBlip r:embed="{svg_rId}"/>
        </a:ext>
      </a:extLst>
    </a:blip>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="0" y="0"/>
      <a:ext cx="{cx_emu}" cy="{cy_emu}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>
"""
    pic_elem = etree.fromstring(pic_xml)

    # Append <p:pic> to the slide shape tree.
    slide.shapes._spTree.append(pic_elem)


# ──────────────────────────────────────────────────────────────────────────────
# CLI
# ──────────────────────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Compile a Beamer .tex file to PDF and export slides as "
                    "vector SVG graphics inside a .pptx presentation."
    )
    p.add_argument("tex_file", type=Path,
                   help="Path to the Beamer .tex source file")
    p.add_argument("--output", "-o", type=Path, default=None,
                   help="Output .pptx path (default: <tex_stem>.pptx alongside the source)")
    p.add_argument("--latex-engine", "-e", default="pdflatex",
                   choices=["pdflatex", "xelatex", "lualatex"],
                   help="LaTeX engine to use (default: pdflatex)")
    p.add_argument("--keep-build", action="store_true",
                   help="Keep the temporary build directory after completion "
                        "(ignored when --build-dir is set, as that dir is never deleted)")
    p.add_argument("--build-dir", "-b", type=Path, default=None,
                   metavar="DIR",
                   help="Use DIR as the build directory instead of a system temp dir. "
                        "The directory is created if it does not exist and is never "
                        "deleted automatically (implies --keep-build).")
    p.add_argument("--pdf", type=Path, default=None,
                   help="Skip compilation and use this existing PDF directly")
    p.add_argument("--latex-cwd", type=Path, default=None,
                   metavar="DIR",
                   help="Working directory for LaTeX/biber/bibtex "
                        "(default: directory containing the .tex file). "
                        "Affects how relative \\input, \\includegraphics, "
                        "and \\bibliography paths in the source resolve.")
    return p.parse_args()


def _run_build(args: argparse.Namespace, build_dir: Path) -> None:
    """Core build logic, shared by the temp-dir and user-defined-dir paths."""
    tex_path: Path = args.tex_file.resolve()
    output_path: Path = args.output or tex_path.with_suffix(".pptx")
    svg_dir = build_dir / "svgs"
    latex_cwd = (args.latex_cwd or tex_path.parent).resolve()

    # ── Compile or use existing PDF ───────────────────────────────────────────
    if args.pdf:
        pdf_path = args.pdf.resolve()
        if not pdf_path.exists():
            sys.exit(f"[error] PDF not found: {pdf_path}")
        print(f"[1/3] Using existing PDF: {pdf_path}")
    else:
        pdf_path = compile_latex(tex_path, args.latex_engine, build_dir, latex_cwd)

    # ── Extract SVGs ──────────────────────────────────────────────────────────
    svg_paths = extract_svgs(pdf_path, svg_dir)

    # ── Build PPTX ────────────────────────────────────────────────────────────
    build_pptx(svg_paths, output_path)


def main() -> None:
    args = parse_args()

    tex_path: Path = args.tex_file.resolve()
    if not tex_path.exists():
        sys.exit(f"[error] File not found: {tex_path}")

    if args.build_dir:
        # User-defined build directory — create it and never auto-delete it.
        build_dir = args.build_dir.resolve()
        build_dir.mkdir(parents=True, exist_ok=True)
        print(f"  Build directory: {build_dir}")
        _run_build(args, build_dir)
    else:
        # Default: auto-managed temp directory.
        with tempfile.TemporaryDirectory(prefix="beamer2pptx_") as tmp:
            build_dir = Path(tmp)
            _run_build(args, build_dir)

            if args.keep_build:
                kept = tex_path.parent / (tex_path.stem + "_build")
                shutil.copytree(tmp, str(kept), dirs_exist_ok=True)
                print(f"  Build directory kept at: {kept}")

    print("\nDone.")


if __name__ == "__main__":
    main()
