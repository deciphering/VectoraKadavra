# VectoraKadavra

**VectoraKadavra** is a powerful Python CLI utility that programmatically converts embedded SVG graphics in Microsoft PowerPoint into native, editable PowerPoint vector shapes (`msoFreeform`).

Unlike PowerPoint's built-in "Convert to Shape" button which discards crops, loses inherited group transformations, and resets orientation, VectoraKadavra ensures 100% pixel-perfect preservation of your original design. 

## Features

- **True Headless Background Execution**: Uses advanced Windows API techniques (ToolWindow styling + off-screen positioning) to run PowerPoint silently in the background without stealing focus or appearing in your Taskbar / Alt-Tab menu.
- **Deep Group Unpacking**: Recursively scans and unpacks deeply nested groups to ensure no embedded SVGs are skipped.
- **Transformation Inheritance**: Automatically ensures SVGs inherit absolute scaling, rotation, and positioning from their parent groups before conversion.
- **Flawless Crop Preservation**: Uses a clever "Duplicate & Uncrop" engine to extract the exact unscaled physical dimensions of heavily cropped or distorted images, preventing squishing or aspect ratio loss.
- **Safe Vector Culling**: Re-aligns bounding boxes to orthogonal axes (0 rotation) to cleanly cull out-of-bounds paths without using destructive boolean geometry (`ShapesIntersect`), perfectly preserving your original strokes and colors.
- **Pixel-Perfect Position & Alignment**: Eliminates clipboard offsets and perfectly restores the original exact bounding box coordinates.
- **Orientation & Flip Preservation**: Restores lost horizontal/vertical flips and rotation angles discarded by PowerPoint's native `SVGEdit` engine.
- **Slide Master & Layout Targeting**: Iterates through all presentation designs, master slides, and layout templates, in addition to standard slides.
- **Optional Auto-Ungrouping**: Can automatically ungroup converted shapes into individual editable vector paths.
- **Safe Execution**: Includes automatic backup creation (`.bak`) and prints a beautiful `Rich` tabular summary of all converted objects.

---

## Setup

Install the required dependencies using pip:
```powershell
pip install -r requirements.txt
```

---

## Usage

### 1. Convert all SVGs in a presentation (Masters, Layouts, and Slides)
```powershell
python convert_svg.py presentation.pptx
```

### 2. Convert only SVGs on Slide Masters and Custom Layouts
```powershell
python convert_svg.py presentation.pptx --target masters
```

### 3. Convert and automatically ungroup into individual paths
```powershell
python convert_svg.py presentation.pptx --ungroup
```

### 4. Save to a separate output file with backup
```powershell
python convert_svg.py input.pptx -o output_converted.pptx --backup
```

---

## Command-Line Arguments

| Argument | Description | Default |
| :--- | :--- | :--- |
| `input` | Path to the PowerPoint presentation (`.pptx`) | *Required* |
| `-o, --output` | Path for the converted output file | Overwrites input |
| `-t, --target` | Scope to target: `all`, `masters`, or `slides` | `all` |
| `--ungroup` | Automatically ungroup converted shapes into sub-shapes | `False` |
| `--backup` | Create a `.bak` backup copy before modifying | `False` |

---

## Technical Notes

- When PowerPoint imports an SVG, it assigns it the `msoGraphic` shape type (ID: `31`).
- `convert_svg.py` enters Slide Master View (`ppViewSlideMaster`), selects each SVG graphic on the master or custom layout, and issues the `SVGEdit` MSO command.
- Requirements: Windows OS with Microsoft PowerPoint (Office 365 or 2019+) installed.
