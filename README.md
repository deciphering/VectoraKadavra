# VectoraKadavra 🪄

Programmatically convert SVG graphics inside Microsoft PowerPoint presentations (`.pptx`) to native editable PowerPoint shapes, with full support for **Slide Masters**, **Custom Layouts**, and standard slides.

## Features

- **Slide Master & Layout Targeting**: Iterates through all presentation designs, master slides, and layout templates to convert embedded SVG graphics.
- **Native Shape Fidelity**: Leverages PowerPoint's native COM engine (`SVGEdit`) to ensure 100% fidelity with the built-in "Convert to Shape" command.
- **Optional Auto-Ungrouping**: Can automatically ungroup converted shapes into individual editable vector paths.
- **Safe Execution**: Includes automatic backup creation (`.bak`) and detailed tabular reporting of all converted objects.

---

## Setup

The Python virtual environment is configured in `.venv/`.

To activate the environment:
```powershell
.\.venv\Scripts\Activate.ps1
```

Or run directly using the venv's Python executable:
```powershell
.\.venv\Scripts\python.exe convert_svg.py <input.pptx>
```

---

## Usage

### 1. Convert all SVGs in a presentation (Masters, Layouts, and Slides)
```powershell
.\.venv\Scripts\python.exe convert_svg.py presentation.pptx
```

### 2. Convert only SVGs on Slide Masters and Custom Layouts
```powershell
.\.venv\Scripts\python.exe convert_svg.py presentation.pptx --target masters
```

### 3. Convert and automatically ungroup into individual paths
```powershell
.\.venv\Scripts\python.exe convert_svg.py presentation.pptx --ungroup
```

### 4. Save to a separate output file with backup
```powershell
.\.venv\Scripts\python.exe convert_svg.py input.pptx -o output_converted.pptx --backup
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
