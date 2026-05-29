# beampptx

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![GitHub stars](https://img.shields.io/github/stars/kocurvik/beampptx)](https://github.com/kocurvik/beampptx/stargazers)
[![Python Version](https://img.shields.io/badge/python-3.6+-blue.svg)](https://www.python.org/downloads/)

Convert LaTeX Beamer slides to PowerPoint presentations with **flawless vector graphics** and **embedded video support**. Mostly vibe-coded using claude-cli and gemini-cli.

## Why beampptx?

Converting Beamer slides to PowerPoint often results in blurry images or lost functionality. `beampptx` solves this by:
- **Vector Fidelity**: Every slide is embedded as a full-bleed SVG vector graphic, ensuring perfect sharpness at any zoom level.
- **Dynamic Overlays**: Supports Beamer transitions (`\pause`, `\alt`, `<1->`, etc.) by expanding them into individual static slides.
- **Native Video**: Automatically extracts videos included via the `movie15` package's `\includemovie` or the `multimedia` package's `\movie` command and embeds them as native PowerPoint video shapes with support for autoplay, looping, and volume.
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
Note that for this option, the videos will not be embedded correctly, but the rest of the content should be exported as vector graphics.


### Advanced Options
```bash
# Specify output filename
beampptx presentation.tex -o final_talk.pptx

# Use a different LaTeX engine
beampptx presentation.tex --latex-engine xelatex

# Keep the temporary build files for debugging
beampptx presentation.tex --keep-build
```

## Examples

The repository includes a `test/` folder with several examples demonstrating different features. To try them out, clone the repository and run:

```bash
# Basic features (math, lists, etc.)
beampptx test/example.tex

# Overlays and transitions (\pause, \alt, etc.)
beampptx test/example_transitions.tex

# Image inclusion (PNG, JPEG, PDF)
beampptx test/example_images.tex

# Bibliography support (biblatex/biber)
beampptx test/example_bib.tex

# Video with the movie15 package (\includemovie)
beampptx test/example_video.tex

# Multiple videos on a single slide
beampptx test/example_multi_video.tex

# Video with the multimedia package (\movie)
beampptx test/example_movie.tex

# Multiple videos on the same slide with the multimedia package (\movie)
beampptx test/example_multimedia.tex
```

## Features in Detail

### Beamer Overlays
`beampptx` detects frames with multiple slides (e.g., from `\pause` or `<1->`) and creates a separate PowerPoint slide for each state. This preserves the feeling of "animations" when clicking through the presentation.

### Videos
Two LaTeX packages are supported for video inclusion — pick whichever your
source already uses.  In both cases `beampptx` finds the call in your `.tex`
source, locates the corresponding annotation in the compiled PDF, and
embeds the video file directly into the `.pptx` as a native PowerPoint
movie shape.

**`movie15` — `\includemovie`:**

```latex
\usepackage{movie15}
...
\includemovie[autoplay, poster=image.png, repeat, volume=0.5]
              {width}{height}{video.mp4}
```

**`multimedia` — `\movie` (ships with Beamer):**

```latex
\usepackage{multimedia}
...
\movie[width=8cm, height=4.5cm, autostart, loop]
      {\includegraphics[width=8cm]{poster.png}}
      {video.mp4}
```

Recognised options:

| Behaviour          | `movie15`               | `multimedia`          |
| ------------------ | ----------------------- | --------------------- |
| Auto-play          | `autoplay`              | `autostart`           |
| Loop               | `repeat` / `palindrome` | `loop` / `palindrome` |
| Initial volume     | `volume=0.0..1.0`       | *(not supported by the package)* |
| Poster image       | `poster=file`           | first arg (`\includegraphics{file}`) |

Note that currently, using pdfs for poster images is not supported. Likewise, the autoplay option requires activating the "next" slide for the videos to run. The palindrome option will just loop the video normally instead of reversing it for even playtroughs.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
