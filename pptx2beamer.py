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
import re

def sanitize_for_latex(text):
    """Remove characters that are invalid for LaTeX command names."""
    return re.sub(r'[^a-zA-Z0-9]', '', text)

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

    # Extract colors - both scheme and custom colors
    colors = {}
    
    # Standard theme colors
    color_scheme = root.find('.//a:clrScheme', ns)
    if color_scheme:
        for color_element in color_scheme:
            tag_name = color_element.tag.split('}')[-1]
            srgb_color = color_element.find('a:srgbClr', ns)
            sys_color = color_element.find('a:sysClr', ns)
            if srgb_color is not None:
                colors[tag_name] = srgb_color.get('val')
            elif sys_color is not None:
                # Handle system colors like window color (usually white)
                last_clr = sys_color.get('lastClr', 'FFFFFF')
                colors[tag_name] = last_clr
        
        # Add standard background and text color mappings if they don't exist
        if 'bg1' not in colors and 'lt1' in colors:
            colors['bg1'] = colors['lt1']  # Background 1 = Light 1 (typically white)
        if 'bg2' not in colors and 'dk1' in colors:
            colors['bg2'] = colors['dk1']  # Background 2 = Dark 1 (typically black)
        if 'tx1' not in colors and 'dk1' in colors:
            colors['tx1'] = colors['dk1']  # Text 1 = Dark 1 (dark text on light background)
        if 'tx2' not in colors and 'dk2' in colors:
            colors['tx2'] = colors['dk2']  # Text 2 = Dark 2
    
    # Custom colors (often defined in theme extras)
    custom_colors = root.findall('.//a:srgbClr', ns)
    for i, custom_color in enumerate(custom_colors):
        val = custom_color.get('val')
        if val and val not in colors.values():
            colors[f'custom{i+1}'] = val

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

def extract_fonts_from_slides(ppt_dir):
    """Extracts font information from actual slide content."""
    fonts_found = set()
    ns = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }
    
    # Search in slides, masters, and layouts
    search_dirs = [
        ppt_dir / 'ppt' / 'slides',
        ppt_dir / 'ppt' / 'slideMasters', 
        ppt_dir / 'ppt' / 'slideLayouts'
    ]
    
    for search_dir in search_dirs:
        if not search_dir.exists():
            continue
            
        for xml_file in search_dir.glob('*.xml'):
            try:
                tree = ET.parse(xml_file)
                root = tree.getroot()
                
                # Look for font references in text runs
                font_elements = root.findall('.//a:latin', ns)
                for font_elem in font_elements:
                    typeface = font_elem.get('typeface')
                    if typeface:
                        fonts_found.add(typeface)
                        
            except ET.ParseError:
                continue
    
    return sorted(list(fonts_found))

def parse_slide_master_styling(ppt_dir):
    """Extracts title/footer styling information from slide masters."""
    styling_info = {
        'has_footer_elements': False,
        'title_positioning': 'default',
        'footer_elements': []
    }
    
    ns = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }
    
    masters_dir = ppt_dir / 'ppt' / 'slideMasters'
    if not masters_dir.exists():
        return styling_info
    
    for master_file in masters_dir.glob('*.xml'):
        try:
            tree = ET.parse(master_file)
            root = tree.getroot()
            
            # Look for footer elements (rectangles, logos, etc.)
            shapes = root.findall('.//p:sp', ns) + root.findall('.//p:pic', ns)
            
            for shape in shapes:
                # Check position - if in bottom area of slide, likely footer
                xfrm = shape.find('.//a:xfrm', ns)
                if xfrm is not None:
                    off = xfrm.find('a:off', ns)
                    if off is not None:
                        y = int(off.get('y', '0'))
                        # If positioned in bottom 20% of slide (>5.5M EMU for standard slide)
                        if y > 5500000:
                            styling_info['has_footer_elements'] = True
                            
                            # Identify element type
                            cNvPr = shape.find('.//p:cNvPr', ns)
                            if cNvPr is not None:
                                name = cNvPr.get('name', '').lower()
                                if 'rectangle' in name:
                                    styling_info['footer_elements'].append('background_rect')
                                elif 'picture' in name or 'logo' in name:
                                    styling_info['footer_elements'].append('logo')
            
        except ET.ParseError:
            continue
    
    return styling_info

def parse_slide_layouts(ppt_dir, theme_colors):
    """Parses all slide layouts for their specific styling."""
    layouts = {}
    ns = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }

    layout_dir = ppt_dir / 'ppt' / 'slideLayouts'
    if not layout_dir.exists():
        return layouts

    for layout_file in layout_dir.glob('*.xml'):
        try:
            tree = ET.parse(layout_file)
            root = tree.getroot()
            
            layout_name = root.find('.//p:cSld', ns).get('name')
            layouts[layout_name] = {
                'name': layout_name,
                'color_overrides': {},
                'placeholders': {},
                'background_color': None,
                'background_image': None
            }

            # Parse color overrides
            color_map_override = root.find('.//p:clrMapOvr/a:overrideClrMapping', ns)
            if color_map_override is not None:
                for key, value in color_map_override.attrib.items():
                    layouts[layout_name]['color_overrides'][key] = value

            # Find solid background color
            bg_element = root.find('.//p:bg/p:bgPr/a:solidFill/a:schemeClr', ns)
            if bg_element is not None:
                bg_color = bg_element.get('val')
                layouts[layout_name]['background_color'] = bg_color

            # Find placeholders with detailed positioning and styling
            for sp in root.findall('.//p:sp', ns):
                ph = sp.find('.//p:nvPr/p:ph', ns)
                if ph is not None:
                    ph_type = ph.get('type', 'body')
                    ph_idx = ph.get('idx', '0')
                    
                    # Create unique placeholder key for multiple placeholders of same type
                    placeholder_key = f"{ph_type}_{ph_idx}" if ph_idx != '0' else ph_type
                    
                    # Extract positioning information from xfrm element
                    position = {'x': 0, 'y': 0, 'width': 0, 'height': 0}
                    xfrm = sp.find('.//a:xfrm', ns)
                    if xfrm is not None:
                        off = xfrm.find('.//a:off', ns)
                        ext = xfrm.find('.//a:ext', ns)
                        if off is not None:
                            position['x'] = int(off.get('x', 0))
                            position['y'] = int(off.get('y', 0))
                        if ext is not None:
                            position['width'] = int(ext.get('cx', 0))
                            position['height'] = int(ext.get('cy', 0))
                    
                    # Extract styling information
                    styling = {
                        'color': None,
                        'font_size': None,
                        'bold': False,
                        'alignment': 'left',
                        'anchor': 'top'
                    }
                    
                    # Look for color in different places within the placeholder
                    color_sources = [
                        './/a:solidFill/a:schemeClr',
                        './/a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr',
                        './/a:defRPr/a:solidFill/a:schemeClr',
                        './/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr'
                    ]
                    
                    for color_path in color_sources:
                        color_element = sp.find(color_path, ns)
                        if color_element is not None:
                            styling['color'] = color_element.get('val')
                            break
                    
                    # Extract font size
                    font_size_element = sp.find('.//a:lvl1pPr/a:defRPr', ns) or sp.find('.//a:defRPr', ns)
                    if font_size_element is not None:
                        sz = font_size_element.get('sz')
                        if sz:
                            styling['font_size'] = int(sz) / 100  # Convert from PowerPoint units to points
                        styling['bold'] = font_size_element.get('b') == '1'
                    
                    # Extract alignment
                    alignment_element = sp.find('.//a:lvl1pPr', ns)
                    if alignment_element is not None:
                        algn = alignment_element.get('algn', 'l')
                        alignment_map = {'l': 'left', 'r': 'right', 'ctr': 'center', 'j': 'justify'}
                        styling['alignment'] = alignment_map.get(algn, 'left')
                    
                    # Extract anchor (vertical alignment)
                    anchor_element = sp.find('.//a:bodyPr', ns)
                    if anchor_element is not None:
                        anchor = anchor_element.get('anchor', 't')
                        anchor_map = {'t': 'top', 'b': 'bottom', 'ctr': 'center'}
                        styling['anchor'] = anchor_map.get(anchor, 'top')
                    
                    # Store both color (for backward compatibility) and full styling info
                    layouts[layout_name]['placeholders'][placeholder_key] = styling['color']
                    
                    # Store detailed placeholder information in a new structure
                    if 'detailed_placeholders' not in layouts[layout_name]:
                        layouts[layout_name]['detailed_placeholders'] = {}
                    
                    layouts[layout_name]['detailed_placeholders'][placeholder_key] = {
                        'type': ph_type,
                        'index': ph_idx,
                        'position': position,
                        'styling': styling
                    }

            # Find background image for this layout
            rels_file = layout_dir / '_rels' / f'{layout_file.name}.rels'
            if rels_file.exists():
                image_name = find_background_image_in_xml(layout_file, rels_file, ns)
                if image_name:
                    layouts[layout_name]['background_image'] = image_name

        except (ET.ParseError, AttributeError):
            continue
            
    return layouts


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
            ("framesubtitle", "accent1", ""),
            ("title", "dk1", "lt1"),
            ("block title", "lt1", "accent2"),
            ("block body", "black", "dk1"),
        ]

        for element, fg_color, bg_color in color_mappings:
            fg = f"ppt{fg_color}" if fg_color in available_colors else "black"
            
            if bg_color == "":
                # No background color specified
                f.write(f"\\setbeamercolor{{{element}}}{{fg={fg}}}\n")
            else:
                bg = f"ppt{bg_color}" if bg_color in available_colors else "white"
                f.write(f"\\setbeamercolor{{{element}}}{{fg={fg},bg={bg}}}\n")

        f.write("\n" + r"\mode<all>")

def generate_font_theme(theme_dir, theme_name, fonts, slide_fonts=None):
    """Generates the beamerfonttheme file, respecting major and minor fonts."""
    filepath = theme_dir / f"beamerfonttheme{theme_name}.sty"

    font_substitutions = {
        'Calibri': 'Helvetica', 'Arial': 'Helvetica', 'Segoe UI': 'Helvetica', 'Tahoma': 'Helvetica',
        'GT America': 'Helvetica', 'Avenir': 'Helvetica', 'Proxima Nova': 'Helvetica', 'Montserrat': 'Helvetica',
        'Open Sans': 'Helvetica', 'Source Sans Pro': 'Helvetica', 'Roboto': 'Helvetica', 'Lato': 'Helvetica',
        'Times New Roman': 'Times', 'Cambria': 'Times'
    }

    def get_compatible_font(font_name):
        if not font_name: return "Helvetica", None
        for ppt_font, tex_font in font_substitutions.items():
            if ppt_font.lower() in font_name.lower():
                return tex_font, ppt_font
        return font_name, None

    with open(filepath, 'w') as f:
        f.write(f"% Font theme for {theme_name}\n")
        f.write(r"\mode<presentation>" + "\n\n")
        f.write("% Requires XeLaTeX or LuaLaTeX for font support\n")
        f.write(r"\RequirePackage{fontspec}" + "\n\n")

        major_font_name = fonts.get('major', 'Times New Roman')
        minor_font_name = fonts.get('minor', 'Arial')

        # If slide fonts are detected, they can override the theme minor font
        if slide_fonts:
            # A simple heuristic: prefer sans-serif fonts found on slides for body text
            sans_serif_candidates = ['Arial', 'Helvetica', 'Calibri', 'Segoe UI', 'Avenir', 'Proxima Nova']
            for cand in sans_serif_candidates:
                for sf in slide_fonts:
                    if cand.lower() in sf.lower():
                        minor_font_name = sf
                        break
                if minor_font_name != fonts.get('minor', 'Arial'): break
        
        major_font, orig_major = get_compatible_font(major_font_name)
        minor_font, orig_minor = get_compatible_font(minor_font_name)

        f.write(f"% Theme fonts: major='{major_font_name}', minor='{minor_font_name}'\n")
        if slide_fonts:
            f.write(f"% Fonts found in slides: {', '.join(slide_fonts)}\n")
        
        f.write("\n% --- Font Definitions ---\n")
        f.write(f"% Body font (minor font): '{orig_minor or minor_font_name}' -> Using: '{minor_font}'\n")
        f.write(f"\\setsansfont{{{minor_font}}}[Ligatures=TeX]\n")
        f.write(f"\\setmainfont{{{minor_font}}}[Ligatures=TeX] % Default to sans-serif for main text\n\n")

        f.write(f"% Title font (major font): '{orig_major or major_font_name}' -> Using: '{major_font}'\n")
        f.write(f"\\newfontfamily\\titlefont{{{major_font}}}[Ligatures=TeX]\n\n")

        f.write("% --- Beamer Font Assignments ---\n")
        f.write(r"\setbeamerfont{normal text}{size=\normalsize}" + "\n")
        f.write(r"\setbeamerfont{title}{family=\titlefont, size=\huge, series=\bfseries}" + "\n")
        f.write(r"\setbeamerfont{frametitle}{family=\titlefont, size=\Large, series=\bfseries}" + "\n")
        f.write(r"\setbeamerfont{framesubtitle}{family=\titlefont, size=\normalsize, series=\mdseries}" + "\n")
        f.write(r"\setbeamerfont{subtitle}{family=\titlefont, size=\large, series=\mdseries}" + "\n")
        f.write(r"\setbeamerfont{author}{family=\titlefont, size=\normalsize}" + "\n")
        f.write(r"\setbeamerfont{institute}{family=\titlefont, size=\small}" + "\n")
        f.write(r"\setbeamerfont{date}{family=\titlefont, size=\small}" + "\n")
        f.write(r"\setbeamerfont{block title}{size=\normalsize, series=\bfseries}" + "\n")

        f.write("\n" + r"\mode<all>")

def convert_ppt_to_beamer_position(position, paper_width=12192000, paper_height=6858000):
    """Convert PowerPoint coordinates to LaTeX/Beamer positioning.
    
    PowerPoint uses EMU (English Metric Units) where:
    - 914400 EMU = 1 inch
    - Standard slide is 12192000 x 6858000 EMU (10" x 7.5")
    """
    # Convert EMU to relative positioning (0.0 to 1.0)
    x_rel = position['x'] / paper_width
    y_rel = position['y'] / paper_height
    width_rel = position['width'] / paper_width
    height_rel = position['height'] / paper_height
    
    # Convert to LaTeX units (approximations for typical slide dimensions)
    x_cm = x_rel * 25.4  # Approximate slide width in cm
    y_cm = y_rel * 19.05  # Approximate slide height in cm  
    width_cm = width_rel * 25.4
    height_cm = height_rel * 19.05
    
    return {
        'x_rel': x_rel,
        'y_rel': y_rel,
        'width_rel': width_rel,
        'height_rel': height_rel,
        'x_cm': x_cm,
        'y_cm': y_cm,
        'width_cm': width_cm,
        'height_cm': height_cm
    }

def generate_beamer_frame_template(layout_name, detailed_placeholders):
    """Generate custom Beamer templates based on detailed placeholder information."""
    template_lines = []
    
    if not detailed_placeholders:
        return []
    
    # Sort placeholders by vertical position (y coordinate) to render in correct order
    sorted_placeholders = sorted(
        detailed_placeholders.items(),
        key=lambda x: x[1]['position']['y']
    )
    
    # Generate integrated frametitle template (includes subtitle)
    template_lines.append(f"% Custom frametitle template for {layout_name}")
    template_lines.append(f"\\defbeamertemplate{{frametitle}}{{{sanitize_for_latex(layout_name).lower()}}}{{%")
    template_lines.append("  \\begin{tikzpicture}[remember picture, overlay]")
    
    title_found = False
    subtitle_found = False
    
    for placeholder_key, placeholder_info in sorted_placeholders:
        ph_type = placeholder_info['type']
        position = placeholder_info['position']
        styling = placeholder_info['styling']
        
        # Skip if no position information
        if position['width'] == 0 or position['height'] == 0:
            continue
            
        beamer_pos = convert_ppt_to_beamer_position(position)
        
        # Handle title placeholder
        if ph_type == 'title':
            template_lines.append(f"    % Title placeholder at ({beamer_pos['x_rel']:.3f}, {beamer_pos['y_rel']:.3f})")
            template_lines.append(f"    \\node[anchor=north west, text width={beamer_pos['width_rel']:.3f}\\paperwidth] at ([xshift={beamer_pos['x_rel']:.3f}\\paperwidth, yshift=-{beamer_pos['y_rel']:.3f}\\paperheight] current page.north west) {{")
            template_lines.append(f"      \\usebeamerfont{{frametitle}}\\usebeamercolor[fg]{{frametitle}}\\insertframetitle")
            template_lines.append(f"    }};")
            title_found = True
        
        # Handle subtitle placeholder (body_18 with accent1 color) - integrate into frametitle template
        elif ph_type == 'body' and ('18' in placeholder_key or 'subtitle' in placeholder_key.lower()):
            template_lines.append(f"    % Subtitle placeholder at ({beamer_pos['x_rel']:.3f}, {beamer_pos['y_rel']:.3f})")
            template_lines.append(f"    \\ifx\\insertframesubtitle\\@empty\\else")
            template_lines.append(f"    \\node[anchor=north west, text width={beamer_pos['width_rel']:.3f}\\paperwidth] at ([xshift={beamer_pos['x_rel']:.3f}\\paperwidth, yshift=-{beamer_pos['y_rel']:.3f}\\paperheight] current page.north west) {{")
            template_lines.append(f"      \\usebeamerfont{{framesubtitle}}\\usebeamercolor[fg]{{framesubtitle}}\\insertframesubtitle")
            template_lines.append(f"    }};")
            template_lines.append(f"    \\fi")
            subtitle_found = True
    
    template_lines.append("  \\end{tikzpicture}")
    template_lines.append("}")
    template_lines.append("")
    
    # Generate framesubtitle template (now integrated above)
    template_lines.append(f"% Custom framesubtitle template for {layout_name} (now integrated above)")
    template_lines.append(f"\\defbeamertemplate{{framesubtitle}}{{{sanitize_for_latex(layout_name).lower()}}}{{%")
    template_lines.append("  % This template is now integrated into the frametitle template above")
    template_lines.append("}")
    template_lines.append("")
    
    return template_lines

def generate_outer_theme(theme_dir, theme_name, layouts, styling_info):
    """Generates the beameroutertheme file with layout-specific environments."""
    filepath = theme_dir / f"beameroutertheme{theme_name}.sty"

    # Check if any layouts need TikZ for detailed positioning
    needs_tikz = any(layout_data.get('detailed_placeholders', {}) for layout_data in layouts.values())
    
    with open(filepath, 'w') as f:
        f.write(f"% Outer theme for {theme_name}\n")
        f.write(r"\mode<presentation>" + "\n\n")
        f.write(r"\usepackage{etoolbox}" + "\n")
        if needs_tikz:
            f.write(r"\RequirePackage{tikz}" + "\n")
            f.write(r"\usetikzlibrary{positioning}" + "\n")
        f.write("\n")
        f.write(r"% Remove navigation symbols" + "\n")
        f.write(r"\setbeamertemplate{navigation symbols}{}" + "\n\n")

        # Default frametitle
        f.write("% Default Frame title\n")
        f.write(r"\setbeamertemplate{frametitle}{%" + "\n")
        f.write(r"  \vspace{0.5cm}" + "\n")
        f.write(r"  \hspace{1em}{\usebeamerfont{frametitle}\insertframetitle}" + "\n")
        f.write(r"  \vspace{0.2cm}" + "\n")
        f.write(r"}" + "\n\n")

        # Layout environments
        f.write("% --- Slide Layout Environments ---\n")
        for layout_name, layout_data in layouts.items():
            env_name = sanitize_for_latex(layout_data['name']).lower()
            f.write(f"\\newenvironment{{{env_name}}}{{%" + "\n")
            
            # ACTIVATE custom templates if they exist
            detailed_placeholders = layout_data.get('detailed_placeholders')
            if detailed_placeholders:
                f.write(f"  % ACTIVATE the correct templates for this layout\n")
                f.write(f"  \\setbeamertemplate{{frametitle}}[{env_name}]\n")
                f.write(f"  \\setbeamertemplate{{framesubtitle}}[{env_name}]\n")
                f.write(f"  % Original code follows\n")
            
            # Apply solid background color only if no background image
            if layout_data['background_color'] and not layout_data['background_image']:
                bg_color = layout_data['background_color']
                # Resolve through color overrides if present
                if layout_data['color_overrides']:
                    bg_color = layout_data['color_overrides'].get(bg_color, bg_color)
                f.write(f"  % Apply solid background color\n")
                f.write(f"  \\setbeamercolor{{background canvas}}{{bg=ppt{bg_color}}}\n")
            
            # Apply color overrides for text (when background image exists)
            if layout_data['color_overrides'] and layout_data['background_image']:
                f.write("  % Apply layout-specific text colors\n")
                tx1_mapped = layout_data['color_overrides'].get('tx1', 'tx1')
                f.write(f"  \\setbeamercolor{{normal text}}{{fg=ppt{tx1_mapped}}}\n")
                
                # Handle tx2 color overrides for body text
                tx2_mapped = layout_data['color_overrides'].get('tx2', 'tx2')
                f.write(f"  \\setbeamercolor{{structure}}{{fg=ppt{tx2_mapped}}}\n")

            # Apply background using picture environment for proper layering
            if layout_data['background_image']:
                img_path = Path(layout_data['background_image'])
                if img_path.suffix.lower() == '.emf':
                    img_path = img_path.with_suffix('.pdf')
                
                if layout_data['background_color']:
                    # Layer transparent image over solid color background
                    bg_color = layout_data['background_color']
                    if layout_data['color_overrides']:
                        bg_color = layout_data['color_overrides'].get(bg_color, bg_color)
                    f.write(f"  % Apply background with colored background behind transparent PNG\n")
                    f.write(f"  \\usebackgroundtemplate{{%\n")
                    f.write(f"    \\begin{{picture}}(0,0)\n")
                    f.write(f"      \\put(0,-\\paperheight){{\\textcolor{{ppt{bg_color}}}{{\\rule{{\\paperwidth}}{{\\paperheight}}}}}}\n")
                    f.write(f"      \\put(0,-\\paperheight){{\\includegraphics[width=\\paperwidth,height=\\paperheight]{{{img_path}}}}}\n")
                    f.write(f"    \\end{{picture}}%\n")
                    f.write(f"  }}\n")
                else:
                    # Just the image without solid background
                    f.write(f"  % Apply background image\n")
                    f.write(f"  \\usebackgroundtemplate{{\\includegraphics[width=\\paperwidth,height=\\paperheight]{{{img_path}}}}}\n")
            elif not layout_data['background_color']:
                # Only clear background if no solid color is set
                f.write("  \\usebackgroundtemplate{}\n")
            
            # Set placeholder-specific colors
            placeholders = layout_data.get('placeholders', {})
            
            # Frametitle color - check master title style first, then placeholders
            title_color = placeholders.get('title')
            if not title_color:
                # If no explicit title color in layout, use master title style (tx2)
                title_color = 'tx2'
            if layout_data['color_overrides']:
                title_color = layout_data['color_overrides'].get(title_color, title_color)
            f.write(f"  \\setbeamercolor{{frametitle}}{{fg=ppt{title_color}}}\n")
            
            # Handle body text color for tx2 elements
            body_color = placeholders.get('body', 'tx2')
            if layout_data['color_overrides']:
                body_color = layout_data['color_overrides'].get(body_color, body_color)
            f.write(f"  \\setbeamercolor{{item}}{{fg=ppt{body_color}}}\n")
            
            # Handle subtitle if present - check for accent1 usage and indexed placeholders
            subtitle_found = False
            for placeholder_key, color in placeholders.items():
                if 'subtitle' in placeholder_key or (color == 'accent1'):
                    subtitle_color = color
                    if layout_data['color_overrides']:
                        subtitle_color = layout_data['color_overrides'].get(subtitle_color, subtitle_color)
                    f.write(f"  \\setbeamercolor{{framesubtitle}}{{fg=ppt{subtitle_color}}}\n")
                    subtitle_found = True
                    break
            
            # If no explicit subtitle found but we have accent1 placeholders, use accent1 for subtitle
            if not subtitle_found:
                for placeholder_key, color in placeholders.items():
                    if color == 'accent1':
                        f.write(f"  \\setbeamercolor{{framesubtitle}}{{fg=pptaccent1}}\n")
                        break

            # Add custom positioning for placeholders
            if placeholders:
                for placeholder_type, color in placeholders.items():
                    if placeholder_type.startswith('placeholder_'):
                        # Add placeholder-specific positioning
                        placeholder_num = placeholder_type.split('_')[1]
                        f.write(f"  % Custom positioning for {placeholder_type}\n")
                        f.write(f"  \setbeamercolor{{{placeholder_type}}}{{fg=ppt{color}}}\n")
                
            f.write("}{%" + "\n")
            f.write("  % End of layout environment\n")
            f.write("}\n\n")
            
            # Custom frame templates for detailed positioning
            detailed_placeholders = layout_data.get('detailed_placeholders')
            if detailed_placeholders:
                frame_template_lines = generate_beamer_frame_template(layout_name, detailed_placeholders)
                for line in frame_template_lines:
                    f.write(line + "\n")
                f.write("\n")

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

def generate_example_file(theme_dir, theme_name, layouts, media_files):
    """Generates an example .tex file demonstrating the layouts."""
    filepath = theme_dir / "example.tex"

    with open(filepath, 'w') as f:
        f.write("% Example presentation using the generated theme\n")
        f.write("% Compile with XeLaTeX or LuaLaTeX for custom fonts\n\n")
        f.write("% !TEX TS-program = lualatex\n\n")
        f.write(r"\documentclass[11pt,aspectratio=169]{beamer}" + "\n")
        f.write(r"\usepackage{graphicx}" + "\n")
        f.write(r"\usepackage{tikz}" + "\n\n")
        f.write(f"\\usetheme{{{theme_name}}}\n\n")
        f.write(r"\title{Sample Presentation}" + "\n")
        f.write(r"\subtitle{Generated from PowerPoint Template}" + "\n")
        f.write(r"\author{Your Name}" + "\n")
        f.write(r"\institute{Your Institution}" + "\n")
        f.write(r"\date{\today}" + "\n\n")
        f.write(r"\begin{document}" + "\n\n")

        # Demonstrate each layout
        for layout_name, layout_data in layouts.items():
            env_name = sanitize_for_latex(layout_data['name']).lower()
            f.write(f"% Frame using the '{layout_name}' layout\n")
            f.write(f"\\begin{{{env_name}}}\n")
            f.write(f"  \\begin{{frame}}\n")
            f.write(f"    \\frametitle{{{layout_name} Layout}}\n")
            
            # Add subtitle for layouts that have subtitle placeholders
            if '1 Column - Subhead' in layout_name:
                f.write(f"    \\framesubtitle{{This is the blue subtitle text}}\n")
            elif 'Executive Summary' in layout_name:
                f.write(f"    \\framesubtitle{{Key Takeaways for Q3 2024}}\n")
            
            f.write(f"    This frame uses the \\texttt{{{env_name}}} environment.\n")
            if layout_data['background_image']:
                f.write(f"    It includes the background image: {layout_data['background_image']}\n")
            f.write(r"  \end{frame}" + "\n")
            f.write(f"\\end{{{env_name}}}\n\n")

        f.write(r"\end{document}" + "\n")

def generate_conversion_report(output_dir, colors, fonts, slide_fonts, layouts, styling_info):
    """Generates a conversion report with notes about visual fidelity."""
    filepath = output_dir / "CONVERSION_NOTES.md"
    
    with open(filepath, 'w') as f:
        f.write("# PowerPoint to Beamer Conversion Report\n\n")
        f.write("## What Was Successfully Converted\n\n")
        
        # Colors
        f.write(f"### Colors ({len(colors)} found)\n")
        if colors:
            for name, hex_val in list(colors.items())[:10]:  # Show first 10
                f.write(f"- `{name}`: #{hex_val}\n")
            if len(colors) > 10:
                f.write(f"- ...and {len(colors) - 10} more colors\n")
        else:
            f.write("- No colors extracted\n")
        f.write("\n")
        
        # Fonts
        f.write(f"### Fonts\n")
        if fonts:
            f.write(f"- Theme fonts: {fonts}\n")
        if slide_fonts:
            f.write(f"- Fonts used in slides: {', '.join(slide_fonts)}\n")
        if not fonts and not slide_fonts:
            f.write("- No fonts detected\n")
        f.write("\n")
        
        # Layouts
        f.write(f"### Slide Layouts ({len(layouts)} found)\n")
        for layout_name, layout_data in layouts.items():
            f.write(f"- **{layout_name}**\n")
            if layout_data['background_color']:
                f.write(f"  - Background Color: `{layout_data['background_color']}`\n")
            if layout_data['background_image']:
                f.write(f"  - Background Image: `{layout_data['background_image']}`\n")
            if layout_data['placeholders']:
                f.write(f"  - Placeholders: `{layout_data['placeholders']}`\n")
            if layout_data['color_overrides']:
                f.write(f"  - Color Overrides: `{layout_data['color_overrides']}`\n")
        f.write("\n")
        
        # Limitations
        f.write("## Conversion Limitations\n\n")
        f.write("### What Cannot Be Converted Automatically\n")
        f.write("- **Exact positioning**: PowerPoint positioning differs from LaTeX\n")
        f.write("- **Complex vector graphics**: Shapes and drawings are not converted.\n")
        f.write("- **Animations**: PowerPoint animations are not supported in Beamer\n\n")
        
        # Manual adjustments
        f.write("### Suggested Manual Adjustments\n")
        f.write("1. **Review Layouts**: Check the generated environments for each slide layout.\n")
        f.write("2. **Customize Fonts**: Install corporate fonts or adjust substitutions in `beamerfonttheme.sty`.\n")
        f.write("3. **Add Logos**: Manually add company logos using `\\includegraphics`.\n")


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
        slide_fonts = extract_fonts_from_slides(temp_path)
        layouts = parse_slide_layouts(temp_path, colors)
        styling_info = parse_slide_master_styling(temp_path)

        print(f"Found {len(colors)} colors, {len(fonts)} theme fonts, {len(slide_fonts)} slide fonts, {len(layouts)} layouts")
        if slide_fonts:
            print(f"Slide fonts detected: {', '.join(slide_fonts[:5])}" + ("..." if len(slide_fonts) > 5 else ""))
        if styling_info['has_footer_elements']:
            print(f"Detected footer elements: {', '.join(styling_info['footer_elements'])}")

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
        generate_font_theme(output_dir, theme_name, fonts, slide_fonts)
        generate_outer_theme(output_dir, theme_name, layouts, styling_info)
        generate_inner_theme(output_dir, theme_name)
        generate_main_theme_file(output_dir, theme_name)
        generate_example_file(output_dir, theme_name, layouts, media_files)
        
        # Generate conversion report
        generate_conversion_report(output_dir, colors, fonts, slide_fonts, layouts, styling_info)

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