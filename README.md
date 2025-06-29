# pptx2beamer.py: PowerPoint to Beamer Converter

## 1. Purpose

The `pptx2beamer.py` script automatically generates Beamer LaTeX themes from Microsoft PowerPoint (`.pptx`) templates. It performs extraction of visual elements including colors, fonts, and layouts to create themes that closely match the original PowerPoint design, reducing manual setup time for PowerPoint-to-Beamer conversions.

---

## 2. Input

- **Required:**
  - `pptx_file` (Path to the input `.pptx` file)
- **Optional:**
  - `--output-dir` or `-o` (Name of the output directory for the generated theme. Defaults to `beamertheme_{pptx_name}`.)

---

## 3. Output

The script creates a theme directory containing:

- **Enhanced Beamer Theme Files (`.sty`):**
  - `beamercolortheme{themename}.sty`: **Extracts colors** including the complete PowerPoint color scheme, custom brand colors, and theme variations.
  - `beamerfonttheme{themename}.sty`: **Detects fonts** from both theme definitions and actual slide content. Supports corporate fonts like Avenir, Proxima Nova with smart OS-agnostic substitutions. Includes `fontspec` configuration for XeLaTeX/LuaLaTeX.
  - `beamerinnertheme{themename}.sty`: Provides modern inner theme styles (rounded blocks, circle itemize markers).
  - `beameroutertheme{themename}.sty`: **Analyzes** PowerPoint template structure to generate appropriate title and footer styling. Detects footer elements (logos, decorative rectangles) and adapts styling accordingly.
  - `beamertheme{themename}.sty`: The main theme file that coordinates all components.
- **Media Files:** All images (PNG, JPG, EMF, etc.) are extracted and copied to the output directory.
- **Converted Vector Graphics:** If `inkscape` is installed, `.emf` vector images are automatically converted to `.pdf` for LaTeX compatibility.
- **Example Presentation (`example.tex`):** Ready-to-compile demonstration with title page, content frames, and background usage examples.
- **Conversion Report (`CONVERSION_NOTES.md`):** Detailed analysis of what was extracted, limitations, and manual adjustment recommendations.

---

## 4. Core Functionalities

- **Advanced PPTX Analysis:** Deep extraction from the complete `.pptx` archive structure for comprehensive visual element detection.
- **Enhanced XML Parsing:**
  - **Complete Color Extraction:** Parses both standard theme colors and custom color definitions to extract colors including brand-specific palettes.
  - **Multi-Source Font Detection:** Analyzes theme definitions, slide masters, layouts, and actual slide content to identify all fonts in use.
  - **Smart Background Detection:** Identifies background images from multiple sources including explicit fills and positioned graphics.
  - **Layout Intelligence:** Analyzes slide master structure to detect and preserve footer elements, logos, and decorative elements.
- **Professional Image Handling:**
  - Copies all media files with format preservation.
  - Automatic EMF-to-PDF conversion using `inkscape` for vector graphics.
  - Background image optimization and cataloging.
- **Theme Generation:** Creates `.sty` files with approximate styling that adapts to the original template's complexity.
- **Comprehensive Documentation:** Generates conversion reports and usage examples for professional deployment.

---

## 5. Capabilities and Limitations

### ✅ **What the Script Handles Well:**
- **Color Fidelity:** Extracts complete color palettes (30+ colors) including custom brand colors
- **Font Intelligence:** Detects and substitutes corporate fonts (GT America, Avenir, etc.)
- **Multiple Layouts:** Handles complex templates with multiple slide layouts and backgrounds

### ⚠️ **Known Limitations:**
- **Vector Graphics:** Cannot extract custom logos, shapes, or geometric elements (noted in conversion report)
- **Complex Positioning:** PowerPoint's exact positioning may differ from LaTeX/Beamer conventions
- **Animations:** PowerPoint animations are not supported in Beamer
- **Font Licensing:** Corporate fonts may require licensing for distribution
- **Manual Refinement:** Generated themes provide an excellent starting point but may need customization for pixel-perfect matching

### 🔧 **Dependencies:**
- **Required:** Python 3, standard libraries
- **Recommended:** `inkscape` for EMF vector conversion
- **For Fonts:** XeLaTeX or LuaLaTeX for advanced typography support

---

## 6. Usage Examples

### Basic Usage
```bash
python3 pptx2beamer.py "YourTemplate.pptx"
```

---

## 7. Workflow and Next Steps

### 1. **Generate the Theme**
```bash
python3 pptx2beamer.py your_template.pptx
```

### 2. **Review the Conversion Report**
- Open `CONVERSION_NOTES.md` to understand what was extracted
- Check for corporate font recommendations
- Note any manual adjustments needed

### 3. **Test the Generated Theme**
```bash
cd beamertheme_yourtemplate
lualatex example.tex  # Recommended for font support
# or: xelatex example.tex
```

### 4. **Customize as Needed**
- **Colors:** Modify `beamercolortheme*.sty` if brand colors need adjustment
- **Fonts:** Install corporate fonts or adjust substitutions in `beamerfonttheme*.sty`
- **Layouts:** Customize `beameroutertheme*.sty` for logos and positioning
- **Backgrounds:** Use `\usebackground{filename}` commands as documented

### 5. **Professional Deployment**
- Review all generated files for corporate compliance
- Test with your actual presentation content
- Distribute theme files with proper attribution
