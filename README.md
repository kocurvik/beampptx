# beampptx

[![GitHub License](https://img.shields.io/github/license/kocurvik/beampptx)](https://github.com/kocurvik/beampptx/blob/main/LICENSE)
[![GitHub stars](https://img.shields.io/github/stars/kocurvik/beampptx)](https://github.com/kocurvik/beampptx/stargazers)
[![Python Version](https://img.shields.io/badge/python-3.6+-blue.svg)](https://www.python.org/downloads/)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](https://github.com/kocurvik/beampptx/pulls)

Convert LaTeX Beamer slides to PowerPoint presentations with **flawless vector graphics** and **embedded video support**.

## Why beampptx?

Converting Beamer slides to PowerPoint often results in blurry images or lost functionality. `beampptx` solves this by:
- **Vector Fidelity**: Every slide is embedded as a full-bleed SVG vector graphic, ensuring perfect sharpness at any zoom level.
- **Dynamic Overlays**: Supports Beamer transitions (`\pause`, `\alt`, `<1->`, etc.) by expanding them into individual static slides.
- **Native Video**: Automatically extracts videos included via the `movie15` package and embeds them as native PowerPoint video shapes with support for autoplay and looping.
- **Bibliography Support**: Handles complex LaTeX compilation passes (including `biber` and `bibtex`).

## Installation

### From GitHub (Recommended)
You can install `beampptx` directly from the repository without cloning:

```bash
pip install git+https://github.com/kocurvik/beampptx.git
```

### Local Development
If you have the repository cloned:

```bash
pip install .
```

### System Requirements
- **Python**: 3.6+
- **LaTeX**: A working distribution (MiKTeX, TeX Live) with `pdflatex` (default), `xelatex`, or `lualatex`.
- **Tools**: `biber` or `bibtex` if using bibliographies.

## Usage

Once installed, use the `beampptx` command:

### Basic Conversion
```bash
beampptx presentation.tex
```

### Convert from PDF directly
If you already have a compiled PDF:
```bash
beampptx presentation.pdf
```

### Advanced Options
```bash
# Specify output filename
beampptx presentation.tex -o final_talk.pptx

# Use a different LaTeX engine
beampptx presentation.tex --latex-engine xelatex

# Keep the temporary build files for debugging
beampptx presentation.tex --keep-build
```

## Features in Detail

### Beamer Overlays
`beampptx` detects frames with multiple slides (e.g., from `\pause` or `<1->`) and creates a separate PowerPoint slide for each state. This preserves the feeling of "animations" when clicking through the presentation.

### Videos
Use the `movie15` package to include videos in your Beamer source:
```latex
\usepackage{movie15}
...
\includemovie[autoplay, poster=image.png]{width}{height}{video.mp4}
```
`beampptx` will find these in your `.tex` source and embed `video.mp4` directly into the `.pptx` file.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
