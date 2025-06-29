# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with
code in this repository.

## Project Overview

`pptx2beamer` is a Python script that automatically converts Microsoft
PowerPoint (.pptx) templates into Beamer LaTeX themes. The tool performs
extraction of visual elements including comprehensive color palettes,
intelligent font detection, and layout analysis to generate themes that closely
match the original PowerPoint design.

## Key Commands

### Running the Script

``` bash
python3 pptx2beamer.py template.pptx
python3 pptx2beamer.py template.pptx -o mytheme
```

### Testing Generated Themes

``` bash
cd beamertheme_<name>
lualatex example.tex
# or
xelatex example.tex
```

## Enhanced Core Architecture

### Main Script (pptx2beamer.py)

-   **Enhanced XML Parsing Functions** (lines 24-108): Extract comprehensive
    color schemes, theme fonts, and slide-level font usage from PowerPoint XML
    structure.
-   **Font Detection Functions** (lines 72-107): New
    `extract_fonts_from_slides()` function analyzes actual slide content to
    identify fonts in use beyond theme definitions.
-   **Style Analysis Functions** (lines 109-153): Analyze slide master structure
    to detect footer elements and determine appropriate styling.
-   **Image Processing** (lines 232-265): Convert EMF vector images to PDF using
    `inkscape`.
-   **Enhanced LaTeX Generation** (lines 268-626): Create sophisticated Beamer
    theme files with corporate-grade styling and intelligent font handling.
-   **Conversion Reporting** (lines 627-695): Generate detailed
    `CONVERSION_NOTES.md` documenting extraction results and manual adjustment
    recommendations.
-   **Main Processing** (lines 699-797): Enhanced orchestration with improved
    font detection and reporting.

### Enhanced Output Structure

Generated themes include:

-   `beamertheme<name>.sty` - Main theme file coordinating all components
-   `beamercolortheme<name>.sty` - Extracts colors including custom brand
    colors, theme variations, and complete corporate palettes
-   `beamerfonttheme<name>.sty` - Font detection from both theme and slide
    content, supports corporate fonts (GT America, Avenir, etc.) with smart
    substitutions
-   `beameroutertheme<name>.sty` - Layout, backgrounds, intelligent title/footer
    styling based on original template structure
-   `beamerinnertheme<name>.sty` - Block styles, itemize formatting
-   `example.tex` - Demo presentation with background usage examples
-   `CONVERSION_NOTES.md` - Detailed conversion report with extraction summary,
    limitations, and manual adjustment guidance
-   Media files (PNG, JPG, PDF converted from EMF)

### Key Dependencies

-   **Python Standard Library**: xml.etree.ElementTree, zipfile, argparse,
    pathlib, shutil, subprocess
-   **External Tool**: `inkscape` (for EMF→PDF conversion, optional but
    recommended)
-   **LaTeX Requirements**: XeLaTeX or LuaLaTeX (for custom font support)

### PowerPoint Structure Assumptions

-   Standard `.pptx` ZIP structure with `ppt/theme/theme1.xml`
-   Background images in `ppt/slideMasters/` and `ppt/slideLayouts/`
-   Media files in `ppt/media/`
-   Relationship files in `_rels/` directories

### Enhanced Font Handling

The script performs intelligent font detection and substitution:

**Font Detection Strategy:**

1.  Analyzes theme font definitions (`major`, `minor` fonts)
2.  Scans actual slide content for fonts in use
3.  Prioritizes corporate fonts found in slide content
4.  Applies smart OS-agnostic substitutions using `fontspec`

**Expanded Font Support:**

-   Calibri → Helvetica
-   Arial → Helvetica
-   Times New Roman → Times
-   Cambria → Times
-   Segoe UI → Helvetica
-   Tahoma → Helvetica
-   GT America (all variants) → Helvetica
-   Avenir (all variants) → Helvetica/Helvetica Neue
-   Proxima Nova → Helvetica
-   Montserrat, Open Sans, Source Sans Pro, Roboto, Lato → Helvetica

**Corporate Font Intelligence:**

-   Detects and reports corporate font usage in `CONVERSION_NOTES.md`
-   Provides specific guidance for common corporate fonts (e.g., GT America)
-   Suggests alternatives when corporate fonts aren't available

Generated themes require XeLaTeX or LuaLaTeX for proper font rendering.

### Background Detection

Uses sophisticated heuristics to identify background images:

-   Explicit background fills in XML (`<p:bg>`, `<a:blipFill>`)
-   Large images (\>7M × 5M EMU) positioned near the slide origin
-   Searches both slide masters AND slide layouts for comprehensive coverage
-   Generates a `\usebackground{<image_file>}` command in
    `beameroutertheme<name>.sty` for applying backgrounds.
-   Also provides a `\clearbackground` command to remove them.
-   Available background images are listed in comments in the `beameroutertheme`
    and `example.tex` files.

### Color Extraction Enhancement

-   Standard theme color scheme extraction
-   Custom color detection from theme extras and slide content
-   Brand color identification
-   Complete color palette extraction
-   Intelligent color mapping to Beamer elements with fallbacks

### Conversion Reporting

-   Generates `CONVERSION_NOTES.md` with:
    -   Complete summary of extracted elements (colors, fonts, backgrounds)
    -   Identification of conversion limitations
    -   Specific recommendations for manual adjustments
    -   Corporate font guidance and licensing considerations
    -   Professional deployment checklist
