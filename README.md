# pptx2beamer.py: Overview

## 1. Purpose

The `pptx2beamer.py` script automates the creation of a Beamer LaTeX theme skeleton from a Microsoft PowerPoint (`.pptx`) file. It extracts colors, font names, and background images to generate a functional starting point for a custom theme, significantly reducing manual setup time.

---

## 2. Input

- **Required:**
  - `pptx_file` (Path to the input `.pptx` file)
- **Optional:**
  - `--output-dir` or `-o` (Name of the output directory for the generated theme. Defaults to `beamertheme_{pptx_name}`.)

---

## 3. Output

The script creates a new directory containing:

- **Beamer Theme Files (`.sty`):**
  - `beamercolortheme{themename}.sty`: Defines theme colors based on the PowerPoint color scheme.
  - `beamerfonttheme{themename}.sty`: Suggests LaTeX-compatible font substitutions (e.g., Calibri → Helvetica) in comments. It does **not** activate these fonts automatically.
  - `beamerinnertheme{themename}.sty`: Provides basic inner theme styles (e.g., rounded blocks, circle itemize markers).
  - `beameroutertheme{themename}.sty`: Defines the outer layout, including a basic frametitle, footline, and commands for applying background images.
  - `beamertheme{themename}.sty`: The main theme file that loads all the sub-theme components.
- **Media Files:** All images (PNG, JPG, EMF, etc.) are extracted from the `.pptx` and copied to the output directory.
- **Converted EMF Files:** If `inkscape` is installed, `.emf` vector images are automatically converted to `.pdf` for LaTeX compatibility.
- **Example Presentation (`example.tex`):**
  - A title page.
  - A sample content frame.
  - A frame explaining the available commands for applying backgrounds (`\usebackground`, `\clearbackground`).
  - An example slide demonstrating the use of the first background image found.

---

## 4. Core Functionalities

- **PPTX Unzipping:** Extracts the `.pptx` archive into a temporary directory for analysis.
- **XML Parsing:**
  - Parses `ppt/theme/theme1.xml` to extract the color scheme and font names.
  - Parses `ppt/slideMasters/*.xml` and `ppt/slideLayouts/*.xml` to find background images. It detects both explicit background fills and large images positioned as backgrounds.
- **Image Handling:**
  - Copies all media files from `ppt/media/` to the output directory.
  - Converts `.emf` images to `.pdf` using `inkscape` if it is available in the system's PATH.
- **LaTeX Theme Generation:** Writes the five `.sty` theme files based on the extracted data, applying sensible defaults and modern Beamer conventions.
- **Example LaTeX Generation:** Creates a ready-to-compile `example.tex` file that showcases the generated theme and demonstrates how to use its features.

---

## 5. Assumptions and Limitations

- **PPTX Structure:** Assumes a standard `.pptx` file structure.
- **Inkscape Dependency:** `.emf` to `.pdf` conversion is a critical feature for vector graphics and requires `inkscape` to be installed and accessible. If not found, a warning is issued.
- **Font Handling:** The script identifies fonts but does **not** embed or activate them. It provides commented-out suggestions in the font theme file. The user is responsible for configuring fonts in their LaTeX document, typically using `fontspec` with XeLaTeX or LuaLaTeX.
- **Theme Skeleton:** The generated theme is a **starting point**. Manual refinement in the `.sty` files is expected to achieve a pixel-perfect match with the original PowerPoint design.
- **Background Assignment:** The script creates commands to apply backgrounds but does not automatically assign them to specific slide layouts. The user must manually add commands like `\usebackground{image.pdf}` where needed.

---

## 6. Usage

1.  Ensure you have Python 3 installed. For best results, install `inkscape`.
2.  Run the script from your terminal:
    ```sh
    python pptx2beamer.py "YourTemplate.pptx"
    ```
3.  To specify an output directory, use the `-o` flag:
    ```sh
    python pptx2beamer.py "MyCorporateTemplate.pptx" -o beamerthememycorp
    ```

---

## 7. Next Steps for User

1.  Navigate into the generated theme directory:
    ```sh
    cd beamertheme_yourtemplate
    ```
2.  Compile the example presentation. **LuaLaTeX or XeLaTeX is recommended** to support custom fonts.
    ```sh
    lualatex example.tex
    ```
3.  Review the output `example.pdf` to see the initial theme.
4.  Customize the `.sty` files (especially `beameroutertheme...` and `beamerfonttheme...`) to match your requirements.

