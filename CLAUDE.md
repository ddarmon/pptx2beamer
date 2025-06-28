# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

`pptx2beamer` is a Python script that automatically converts Microsoft PowerPoint (.pptx) templates into Beamer LaTeX theme skeletons. The tool extracts visual elements (colors, fonts, background images) from PowerPoint files and generates a complete Beamer theme with an example presentation.

## Key Commands

### Running the Script
```bash
python pptx2beamer.py template.pptx
python pptx2beamer.py template.pptx -o mytheme
```

### Testing Generated Themes
```bash
cd beamertheme_<name>
lualatex example.tex
# or
xelatex example.tex
```

## Core Architecture

### Main Script (pptx2beamer.py)
- **XML Parsing Functions** (lines 24-179): Extract colors, fonts, and background images from PowerPoint XML structure.
- **Image Processing** (lines 183-216): Convert EMF vector images to PDF using `inkscape`.
- **LaTeX Generation** (lines 220-491): Create Beamer theme files (.sty) and an example presentation.
- **Main Processing** (lines 495-598): Orchestrate extraction, parsing, and generation.

### Output Structure
Generated themes include:
- `beamertheme<name>.sty` - Main theme file
- `beamercolortheme<name>.sty` - Color definitions extracted from PowerPoint
- `beamerfonttheme<name>.sty` - Font suggestions with Windows→LaTeX substitutions
- `beameroutertheme<name>.sty` - Layout, backgrounds, frametitle, footline
- `beamerinnertheme<name>.sty` - Block styles, itemize formatting
- `example.tex` - Demo presentation with background usage examples
- Media files (PNG, JPG, PDF converted from EMF)

### Key Dependencies
- **Python Standard Library**: xml.etree.ElementTree, zipfile, argparse, pathlib, shutil, subprocess
- **External Tool**: `inkscape` (for EMF→PDF conversion, optional but recommended)
- **LaTeX Requirements**: XeLaTeX or LuaLaTeX (for custom font support)

### PowerPoint Structure Assumptions
- Standard `.pptx` ZIP structure with `ppt/theme/theme1.xml`
- Background images in `ppt/slideMasters/` and `ppt/slideLayouts/`
- Media files in `ppt/media/`
- Relationship files in `_rels/` directories

### Font Handling
The script suggests font substitutions in the generated `beamerfonttheme<name>.sty` file for cross-platform compatibility. It does not automatically apply them.
- Calibri → Helvetica
- Arial → Helvetica
- Times New Roman → Times
- Cambria → Times
- Segoe UI → Helvetica
- Tahoma → Helvetica

### Background Detection
Uses heuristics to identify background images:
- Explicit background fills in XML (`<p:bg>`, `<a:blipFill>`)
- Large images (>7M × 5M EMU) positioned near the slide origin
- Generates a `\usebackground{<image_file>}` command in `beameroutertheme<name>.sty` for applying backgrounds.
- Also provides a `\clearbackground` command to remove them.
- Available background images are listed in comments in the `beameroutertheme` and `example.tex` files.