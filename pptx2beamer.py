#!/usr/bin/env python3
#
# pptx2beamer.py
# A script to programmatically generate a Beamer template skeleton
# from a PowerPoint .pptx template file.
#
# Usage:
# python pptx2beamer.py YourTemplate.pptx --output-dir beamerthememycompany
#

import sys
import os
import argparse
import zipfile
import shutil
import tempfile
import subprocess
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime

# --- XML Parsing Functions ---

def parse_theme_xml(theme_path):
    """Parses the theme1.xml file for colors and fonts."""
    if not theme_path.exists():
        print(f"Warning: Theme file {theme_path} not found.")
        return {}, {}

    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}

    try:
        tree = ET.parse(theme_path)
        root = tree.getroot()
    except ET.ParseError as e:
        print(f"Warning: Could not parse theme XML: {e}")
        return {}, {}

    # Extract colors
    colors = {}
    color_scheme = root.find('.//a:clrScheme', ns)
    if color_scheme:
        for color_element in color_scheme:
            tag_name = color_element.tag.split('}')[-1]
            srgb_color = color_element.find('a:srgbClr', ns)
            if srgb_color is not None:
                colors[tag_name] = srgb_color.get('val')

    # Extract fonts
    fonts = {}
    font_scheme = root.find('.//a:fontScheme', ns)
    if font_scheme:
        major_font_element = font_scheme.find('.//a:majorFont/a:latin', ns)
        if major_font_element is not None:
            fonts['major'] = major_font_element.get('typeface')

        minor_font_element = font_scheme.find('.//a:minorFont/a:latin', ns)
        if minor_font_element is not None:
            fonts['minor'] = minor_font_element.get('typeface')

    return colors, fonts

def find_background_images(ppt_dir):
    """Finds background images from slide masters and layouts."""
    backgrounds = {}
    ns = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }

    # Search in both masters and layouts
    search_dirs = [
        ('masters', ppt_dir / 'ppt' / 'slideMasters'),
        ('layouts', ppt_dir / 'ppt' / 'slideLayouts')
    ]

    for dir_type, search_dir in search_dirs:
        if not search_dir.exists():
            continue

        for xml_file in search_dir.glob('*.xml'):
            rels_file = search_dir / '_rels' / f'{xml_file.name}.rels'
            if rels_file.exists():
                image_name = find_background_image_in_xml(xml_file, rels_file, ns)
                if image_name and image_name not in backgrounds:
                    cmd_name = f"usebackground{len(backgrounds) + 1}"
                    backgrounds[image_name] = cmd_name

    return backgrounds

def find_background_image_in_xml(xml_path, rels_path, ns):
    """Helper to find background images in a given XML file."""
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()

        # First, look for actual background fills
        background_patterns = [
            './/p:cSld/p:bg//a:blipFill',
            './/p:bgPr//a:blipFill',
            './/p:bg//a:blipFill',
            './/a:bgFillStyleLst//a:blipFill',
        ]

        for pattern in background_patterns:
            blip_fill = root.find(pattern, ns)
            if blip_fill is not None:
                blip = blip_fill.find('a:blip', ns)
                if blip is not None:
                    r_id = blip.get(f'{{{ns["r"]}}}embed')
                    if r_id:
                        image_name = get_image_from_relationship(rels_path, r_id)
                        if image_name:
                            return image_name

        # If no explicit background, look for large pictures that cover the slide
        pictures = root.findall('.//p:pic', ns)

        for pic in pictures:
            # Get the picture dimensions and position
            xfrm = pic.find('.//a:xfrm', ns)
            if xfrm is not None:
                off = xfrm.find('a:off', ns)
                ext = xfrm.find('a:ext', ns)

                if off is not None and ext is not None:
                    x = int(off.get('x', '0'))
                    y = int(off.get('y', '0'))
                    cx = int(ext.get('cx', '0'))
                    cy = int(ext.get('cy', '0'))

                    # Check if this picture is large and positioned like a background
                    is_large = cx > 7000000 and cy > 5000000
                    is_positioned_as_bg = x < 100000 and y < 100000

                    if is_large and is_positioned_as_bg:
                        # This looks like a background image
                        blip_fill = pic.find('.//p:blipFill', ns)
                        if blip_fill is not None:
                            blip = blip_fill.find('a:blip', ns)
                            if blip is not None:
                                r_id = blip.get(f'{{{ns["r"]}}}embed')
                                if r_id:
                                    image_name = get_image_from_relationship(rels_path, r_id)
                                    if image_name:
                                        return image_name

    except ET.ParseError:
        pass
    return None

def get_image_from_relationship(rels_path, r_id):
    """Get image filename from relationship ID."""
    try:
        rels_tree = ET.parse(rels_path)

        # Handle namespace for relationships XML
        rels_ns = {'pkg': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        relationships = rels_tree.findall('.//pkg:Relationship', rels_ns)

        # Fallback to no namespace if the above doesn't work
        if not relationships:
            relationships = rels_tree.findall('.//Relationship')

        for rel in relationships:
            rel_id = rel.get('Id')
            target = rel.get('Target', '')

            if rel_id == r_id:
                # Only return image files
                if target and any(target.lower().endswith(ext) for ext in ['.emf', '.png', '.jpg', '.jpeg', '.svg', '.bmp']):
                    return Path(target).name

    except ET.ParseError:
        pass
    return None

# --- Image Conversion ---

def convert_emf_to_pdf(output_dir):
    """Converts EMF files to PDF using inkscape if available."""
    emf_files = list(output_dir.glob('*.emf'))
    if not emf_files:
        return

    inkscape_path = shutil.which('inkscape')
    if not inkscape_path:
        print("\nWarning: 'inkscape' not found. EMF files will not be converted to PDF.")
        print("Install Inkscape to convert vector images automatically.")
        return

    print("Converting EMF images to PDF...")
    converted_count = 0

    for emf_file in emf_files:
        pdf_file = emf_file.with_suffix('.pdf')
        try:
            result = subprocess.run([
                inkscape_path,
                f'--export-filename={pdf_file}',
                str(emf_file)
            ], check=True, capture_output=True, text=True)

            print(f"  ✓ Converted {emf_file.name} to {pdf_file.name}")
            converted_count += 1
        except subprocess.CalledProcessError as e:
            print(f"  ✗ Failed to convert {emf_file.name}: {e.stderr.strip()}")
        except Exception as e:
            print(f"  ✗ Error converting {emf_file.name}: {e}")

    if converted_count > 0:
        print(f"Successfully converted {converted_count} EMF files.")

# --- LaTeX File Generation Functions ---

def generate_color_theme(theme_dir, theme_name, colors):
    """Generates the beamercolortheme file."""
    filepath = theme_dir / f"beamercolortheme{theme_name}.sty"

    with open(filepath, 'w') as f:
        f.write(f"% Color theme for {theme_name}\n")
        f.write(r"\mode<presentation>" + "\n\n")

        if not colors:
            f.write("% No colors found in PowerPoint theme\n")
            f.write("% Using default Beamer colors\n\n")
            f.write(r"\mode<all>")
            return

        f.write("% Extracted PowerPoint Colors\n")
        for name, hex_val in colors.items():
            if hex_val:  # Ensure hex value exists
                f.write(f"\\definecolor{{ppt{name}}}{{HTML}}{{{hex_val}}}\n")

        f.write("\n% Color assignments (modify as needed)\n")

        # Use available colors more intelligently
        available_colors = list(colors.keys())

        # Default mappings with fallbacks
        color_mappings = [
            ("normal text", "dk1", "lt1"),
            ("structure", "accent1", "dk1"),
            ("frametitle", "lt1", "dk1"),
            ("title", "dk1", "lt1"),
            ("block title", "lt1", "accent2"),
            ("block body", "black", "dk1"),
        ]

        for element, fg_color, bg_color in color_mappings:
            fg = f"ppt{fg_color}" if fg_color in available_colors else "black"
            bg = f"ppt{bg_color}" if bg_color in available_colors else "white"

            f.write(f"\\setbeamercolor{{{element}}}{{fg={fg},bg={bg}}}\n")

        f.write("\n" + r"\mode<all>")

def generate_font_theme(theme_dir, theme_name, fonts):
    """Generates the beamerfonttheme file."""
    filepath = theme_dir / f"beamerfonttheme{theme_name}.sty"

    # Font mapping for better cross-platform compatibility
    font_substitutions = {
        'Calibri': 'Helvetica',
        'Arial': 'Helvetica',
        'Times New Roman': 'Times',
        'Cambria': 'Times',
        'Segoe UI': 'Helvetica',
        'Tahoma': 'Helvetica'
    }

    def get_compatible_font(font_name):
        for windows_font, replacement in font_substitutions.items():
            if windows_font.lower() in font_name.lower():
                return replacement, windows_font
        return font_name, None

    with open(filepath, 'w') as f:
        f.write(f"% Font theme for {theme_name}\n")
        f.write(r"\mode<presentation>" + "\n\n")

        if not fonts:
            f.write("% No fonts found in PowerPoint theme\n")
            f.write("% Using default Beamer fonts\n\n")
            f.write(r"\mode<all>")
            return

        f.write("% Font theme uses default LaTeX fonts for compatibility\n")
        f.write("% To use custom fonts, add fontspec configuration to your document preamble\n\n")

        # Document font suggestions in comments rather than executing font commands
        if 'major' in fonts:
            major_font, original_major = get_compatible_font(fonts['major'])
            if original_major:
                f.write(f"% Original major font: '{original_major}' -> Suggested: '{major_font}'\n")
            else:
                f.write(f"% Suggested major font: '{major_font}'\n")

        if 'minor' in fonts:
            minor_font, original_minor = get_compatible_font(fonts['minor'])
            if original_minor:
                f.write(f"% Original minor font: '{original_minor}' -> Suggested: '{minor_font}'\n")
            else:
                f.write(f"% Suggested minor font: '{minor_font}'\n")

        f.write("\n% Beamer font settings\n")
        f.write(r"\setbeamerfont{normal text}{size=\normalsize}" + "\n")
        f.write(r"\setbeamerfont{title}{size=\huge,series=\bfseries}" + "\n")
        f.write(r"\setbeamerfont{frametitle}{size=\Large,series=\bfseries}" + "\n")
        f.write(r"\setbeamerfont{block title}{size=\normalsize,series=\bfseries}" + "\n")

        f.write("\n" + r"\mode<all>")

def generate_outer_theme(theme_dir, theme_name, backgrounds):
    """Generates the beameroutertheme file."""
    filepath = theme_dir / f"beameroutertheme{theme_name}.sty"

    with open(filepath, 'w') as f:
        f.write(f"% Outer theme for {theme_name}\n")
        f.write(r"\mode<presentation>" + "\n\n")
        f.write(r"% Remove navigation symbols" + "\n")
        f.write(r"\setbeamertemplate{navigation symbols}{}" + "\n\n")

        # Simple background commands
        f.write("% Background commands (use before \\begin{frame})\n")
        f.write("\\newcommand{\\usebackground}[1]{%\n")
        f.write("  \\usebackgroundtemplate{\\includegraphics[width=\\paperwidth,height=\\paperheight]{#1}}%\n")
        f.write("}\n\n")
        f.write("\\newcommand{\\clearbackground}{%\n")
        f.write("  \\usebackgroundtemplate{}%\n")
        f.write("}\n\n")

        if backgrounds:
            f.write("% Available background images:\n")
            for img_file, cmd_name in backgrounds.items():
                img_path = Path(img_file)
                # Use PDF version if EMF was converted
                if img_path.suffix.lower() == '.emf':
                    img_path = img_path.with_suffix('.pdf')
                f.write(f"% \\usebackground{{{img_path}}}\n")
        else:
            f.write("% No background images found\n")

        # Frame title template
        f.write("% Frame title\n")
        f.write(r"\setbeamertemplate{frametitle}{%" + "\n")
        f.write(r"  \nointerlineskip" + "\n")
        f.write(r"  \begin{beamercolorbox}[wd=\paperwidth,ht=2.5ex,dp=1.5ex]{frametitle}%" + "\n")
        f.write(r"    \hspace*{1em}\insertframetitle" + "\n")
        f.write(r"  \end{beamercolorbox}%" + "\n")
        f.write(r"}" + "\n\n")

        # Footline template
        f.write("% Footline\n")
        f.write(r"\setbeamertemplate{footline}{%" + "\n")
        f.write(r"  \begin{beamercolorbox}[wd=\paperwidth,ht=2.5ex,dp=1ex]{frametitle}%" + "\n")
        f.write(r"    \hfill\usebeamerfont{page number in head/foot}%" + "\n")
        f.write(r"    \insertframenumber{} / \inserttotalframenumber\hspace*{1em}%" + "\n")
        f.write(r"  \end{beamercolorbox}%" + "\n")
        f.write(r"}" + "\n\n")

        f.write(r"\mode<all>")

def generate_inner_theme(theme_dir, theme_name):
    """Generates the beamerinnertheme file."""
    filepath = theme_dir / f"beamerinnertheme{theme_name}.sty"

    with open(filepath, 'w') as f:
        f.write(f"% Inner theme for {theme_name}\n")
        f.write(r"\mode<presentation>" + "\n\n")
        f.write("% Customize itemize, blocks, etc.\n\n")
        f.write("% Rounded blocks with shadow\n")
        f.write(r"\setbeamertemplate{blocks}[rounded][shadow=true]" + "\n\n")
        f.write("% Custom itemize items\n")
        f.write(r"\setbeamertemplate{itemize items}[circle]" + "\n\n")
        f.write(r"\mode<all>")

def generate_main_theme_file(theme_dir, theme_name):
    """Generates the main beamertheme file."""
    filepath = theme_dir / f"beamertheme{theme_name}.sty"
    current_date = datetime.now().strftime("%Y/%m/%d")

    with open(filepath, 'w') as f:
        f.write(f"% Main Beamer theme file for {theme_name}\n")
        f.write(r"\NeedsTeXFormat{LaTeX2e}" + "\n")
        f.write(f"\\ProvidesPackage{{beamertheme{theme_name}}}[{current_date} v1.0 {theme_name.title()} Beamer Theme]\n\n")
        f.write(r"\mode<presentation>" + "\n\n")
        f.write(f"\\usecolortheme{{{theme_name}}}\n")
        f.write(f"\\usefonttheme{{{theme_name}}}\n")
        f.write(f"\\useinnertheme{{{theme_name}}}\n")
        f.write(f"\\useoutertheme{{{theme_name}}}\n")
        f.write("\n" + r"\mode<all>")

def generate_example_file(theme_dir, theme_name, backgrounds, media_files):
    """Generates an example .tex file."""
    filepath = theme_dir / "example.tex"

    with open(filepath, 'w') as f:
        f.write("% Example presentation using the generated theme\n")
        f.write("% Compile with XeLaTeX or LuaLaTeX for custom fonts\n\n")
        f.write("% !TEX TS-program = lualatex\n\n")
        f.write(r"\documentclass[11pt]{beamer}" + "\n")
        f.write(r"\usepackage{graphicx}" + "\n")
        f.write(r"\usepackage{tikz}" + "\n\n")
        f.write(f"\\usetheme{{{theme_name}}}\n\n")
        f.write(r"\title{Sample Presentation}" + "\n")
        f.write(r"\subtitle{Generated from PowerPoint Template}" + "\n")
        f.write(r"\author{Your Name}" + "\n")
        f.write(r"\institute{Your Institution}" + "\n")
        f.write(r"\date{\today}" + "\n\n")
        f.write(r"\begin{document}" + "\n\n")

        # Title slide
        f.write("% Title slide\n")
        f.write("\\clearbackground\n")
        f.write(r"\begin{frame}" + "\n")
        f.write(r"  \titlepage" + "\n")
        f.write(r"\end{frame}" + "\n\n")

        # Content frame
        f.write(r"\begin{frame}{About This Theme}" + "\n")
        f.write(r"  This theme was automatically generated from a PowerPoint template." + "\n\n")
        f.write(r"  \begin{itemize}" + "\n")
        f.write(r"    \item Colors extracted from theme" + "\n")
        f.write(r"    \item Fonts identified and substituted" + "\n")
        f.write(r"    \item Background images included" + "\n")
        f.write(r"    \item Media files copied" + "\n")
        f.write(r"  \end{itemize}" + "\n")
        f.write(r"\end{frame}" + "\n\n")

        # Background demonstration
        f.write(r"\begin{frame}{Available Background Commands}" + "\n")
        f.write(r"  Use these commands to apply backgrounds:" + "\n")
        f.write(r"  \begin{itemize}" + "\n")
        f.write(r"    \item \texttt{\textbackslash usebackground\{filename.pdf\}}" + "\n")
        f.write(r"    \item \texttt{\textbackslash clearbackground}" + "\n")
        f.write(r"  \end{itemize}" + "\n")
        if backgrounds:
            f.write(r"  \vspace{1em}" + "\n")
            f.write(r"  Available images:" + "\n")
            f.write(r"  \begin{itemize}" + "\n")
            for img_file, cmd_name in backgrounds.items():
                img_path = Path(img_file)
                if img_path.suffix.lower() == '.emf':
                    img_path = img_path.with_suffix('.pdf')
                f.write(f"    \\item \\texttt{{{img_path}}}\n")
            f.write(r"  \end{itemize}" + "\n")
        f.write(r"\end{frame}" + "\n\n")

        # Example background slide if images exist
        if backgrounds:
            first_img = list(backgrounds.keys())[0]
            img_path = Path(first_img)
            if img_path.suffix.lower() == '.emf':
                img_path = img_path.with_suffix('.pdf')
            f.write("% Example with background\n")
            f.write(f"\\usebackground{{{img_path}}}\n")
            f.write(r"\begin{frame}{Background Example}" + "\n")
            f.write(f"  This slide demonstrates using \\texttt{{\\textbackslash usebackground}} with \\texttt{{{img_path}}}.\n")
            f.write(r"  \vspace{1em}" + "\n")
            f.write(r"  Use \texttt{\textbackslash clearbackground} to remove the background." + "\n")
            f.write(r"\end{frame}" + "\n\n")

        f.write(r"\end{document}" + "\n")

# --- Main Function ---

def main():
    parser = argparse.ArgumentParser(
        description="Convert a PowerPoint .pptx template to a Beamer theme skeleton.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python pptx2beamer.py template.pptx
  python pptx2beamer.py template.pptx -o mytheme
        """
    )
    parser.add_argument("pptx_file", type=Path,
                       help="Path to the input .pptx file")
    parser.add_argument("--output-dir", "-o", type=str, default=None,
                       help="Output directory name (default: beamertheme_<filename>)")

    args = parser.parse_args()

    # Validate input
    if not args.pptx_file.is_file():
        print(f"Error: '{args.pptx_file}' not found.")
        sys.exit(1)

    if args.pptx_file.suffix.lower() != '.pptx':
        print(f"Error: '{args.pptx_file}' is not a .pptx file.")
        sys.exit(1)

    # Determine output directory and theme name
    if args.output_dir:
        output_dir = Path(args.output_dir)
        theme_name = args.output_dir.replace('beamertheme', '').strip('_')
        if not theme_name:
            theme_name = args.pptx_file.stem.lower().replace(' ', '')
    else:
        base_name = args.pptx_file.stem.lower().replace(' ', '')
        output_dir = Path(f"beamertheme_{base_name}")
        theme_name = base_name

    # Clean theme name
    theme_name = ''.join(c for c in theme_name if c.isalnum())
    if not theme_name:
        theme_name = "custom"

    print(f"Processing: {args.pptx_file}")
    print(f"Output directory: {output_dir}")
    print(f"Theme name: {theme_name}")

    # Create output directory
    if output_dir.exists():
        shutil.rmtree(output_dir)
    output_dir.mkdir(parents=True)

    # Process PowerPoint file
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)

        # Extract PPTX
        try:
            with zipfile.ZipFile(args.pptx_file, 'r') as zip_ref:
                zip_ref.extractall(temp_path)
        except (zipfile.BadZipFile, PermissionError) as e:
            print(f"Error: Could not extract '{args.pptx_file}': {e}")
            sys.exit(1)

        # Parse theme data
        theme_xml_path = temp_path / "ppt" / "theme" / "theme1.xml"
        colors, fonts = parse_theme_xml(theme_xml_path)
        backgrounds = find_background_images(temp_path)

        print(f"Found {len(colors)} colors, {len(fonts)} fonts, {len(backgrounds)} backgrounds")

        # Copy media files
        media_files = []
        media_path = temp_path / "ppt" / "media"
        if media_path.exists():
            for media_file in media_path.iterdir():
                if media_file.is_file():
                    shutil.copy2(media_file, output_dir)
                    media_files.append(media_file.name)
            print(f"Copied {len(media_files)} media files")

        # Convert EMF files
        convert_emf_to_pdf(output_dir)

        # Generate theme files
        print("Generating theme files...")
        generate_color_theme(output_dir, theme_name, colors)
        generate_font_theme(output_dir, theme_name, fonts)
        generate_outer_theme(output_dir, theme_name, backgrounds)
        generate_inner_theme(output_dir, theme_name)
        generate_main_theme_file(output_dir, theme_name)
        generate_example_file(output_dir, theme_name, backgrounds, media_files)

    print("\n" + "="*60)
    print("🎉 Beamer Theme Generation Complete!")
    print("="*60)
    print(f"\nTheme '{theme_name}' created in: {output_dir}/")
    print(f"\nNext steps:")
    print(f"1. cd {output_dir}")
    print(f"2. lualatex example.tex")
    print(f"3. Review and customize the .sty files as needed")

    print(f"\nNote: If background images are missing, placeholder backgrounds will be used.")
    print(f"Background commands available: \\usebackground1 through \\usebackground5")

if __name__ == "__main__":
    main()