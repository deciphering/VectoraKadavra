"""
VectoraKadavra: Convert SVG graphics to native PowerPoint shapes programmatically.
Preserves pixel-perfect position, exact orientation, rotation, horizontal/vertical flips,
and relative sub-shape alignments.
Supports Slide Masters, Custom Layouts, and regular Slides.
"""

import os
import sys
import time
import shutil
import argparse
from pathlib import Path
from rich.console import Console
from rich.table import Table
from rich.panel import Panel

console = Console()

# PowerPoint / Office Constants
MSO_GRAPHIC_TYPES = (28, 31)  # SVG / Office Graphic shape types across Office versions
MSO_GROUP = 6                 # Group shape
MSO_FLIP_HORIZONTAL = 0       # msoFlipHorizontal
MSO_FLIP_VERTICAL = 1         # msoFlipVertical
PP_VIEW_SLIDE = 1             # Slide View
PP_VIEW_SLIDE_MASTER = 2      # Slide Master View
PP_VIEW_NORMAL = 9            # Normal View

def get_powerpoint_app():
    """Initializes and returns the PowerPoint COM Application instance."""
    try:
        import win32com.client
    except ImportError:
        console.print("[bold red]Error:[/bold red] pywin32 is not installed. Run 'pip install pywin32' first.")
        sys.exit(1)

    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        ppt_app.Visible = True
        return ppt_app
    except Exception as e:
        console.print(f"[bold red]Failed to launch Microsoft PowerPoint COM Automation:[/bold red] {e}")
        console.print("[yellow]Make sure Microsoft PowerPoint (Office 365 or 2019+) is installed on this machine.[/yellow]")
        sys.exit(1)

def safe_clipboard_action(action_func, max_retries=5, delay=0.08):
    """Executes a COM clipboard action with retries to handle Windows clipboard lock timing."""
    for attempt in range(max_retries):
        try:
            return action_func()
        except Exception as e:
            if attempt == max_retries - 1:
                raise e
            time.sleep(delay * (attempt + 1))

class SVGToShapeConverter:
    def __init__(self, ppt_app, presentation, ungroup=False, preserve_orientation=True, delay=0.05):
        self.ppt_app = ppt_app
        self.pres = presentation
        self.ungroup = ungroup
        self.preserve_orientation = preserve_orientation
        self.delay = delay
        self.converted_count = 0
        self.errors = []
        self.conversion_log = []
        self.scratch_slide = None

    def _get_or_create_scratch_slide(self):
        """Creates a dedicated scratch slide in Normal view for reliable conversion."""
        if self.scratch_slide is None:
            self.ppt_app.ActiveWindow.ViewType = PP_VIEW_NORMAL
            time.sleep(0.1)
            blank_layout = self.pres.SlideMaster.CustomLayouts.Item(7)
            self.scratch_slide = self.pres.Slides.AddSlide(self.pres.Slides.Count + 1, blank_layout)
        return self.scratch_slide

    def _get_center_coordinates(self, shape_or_range):
        """Calculates the exact visual center coordinates for a Shape or ShapeRange."""
        try:
            if hasattr(shape_or_range, "Count") and shape_or_range.Count > 1:
                items = [shape_or_range.Item(i) for i in range(1, shape_or_range.Count + 1)]
                min_x = min(item.Left for item in items)
                max_x = max(item.Left + item.Width for item in items)
                min_y = min(item.Top for item in items)
                max_y = max(item.Top + item.Height for item in items)
                return ((min_x + max_x) / 2.0, (min_y + max_y) / 2.0)
            else:
                item = shape_or_range.Item(1) if hasattr(shape_or_range, "Count") else shape_or_range
                return (item.Left + item.Width / 2.0, item.Top + item.Height / 2.0)
        except Exception:
            return (0.0, 0.0)

    def _convert_shape(self, shp, target_container, container_name):
        """
        Converts an SVG shape to native PowerPoint shapes and restores all
        pixel-perfect position, rotation, horizontal/vertical flips, and bounding box positions.
        """
        try:
            shape_name = shp.Name
        except Exception:
            shape_name = "Unknown Graphic"

        # 1. Capture original transformation state and center point
        orig_rot = getattr(shp, 'Rotation', 0.0)
        orig_hflip = (getattr(shp, 'HorizontalFlip', 0) == -1)
        orig_vflip = (getattr(shp, 'VerticalFlip', 0) == -1)
        orig_left = getattr(shp, 'Left', 0.0)
        orig_top = getattr(shp, 'Top', 0.0)
        orig_width = getattr(shp, 'Width', 0.0)
        orig_height = getattr(shp, 'Height', 0.0)
        orig_center_x = orig_left + orig_width / 2.0
        orig_center_y = orig_top + orig_height / 2.0

        scratch = self._get_or_create_scratch_slide()

        try:
            # 2. Copy shape to scratch slide
            safe_clipboard_action(lambda: shp.Copy())
            scratch.Select()
            time.sleep(self.delay)
            
            pasted_range = safe_clipboard_action(lambda: scratch.Shapes.Paste())
            pasted_shp = pasted_range.Item(1)

            # 3. Select and trigger Convert to Shape
            pasted_shp.Select()
            time.sleep(self.delay)
            self.ppt_app.CommandBars.ExecuteMso("SVGEdit")
            time.sleep(self.delay)

            sel = self.ppt_app.ActiveWindow.Selection
            if sel.Type != 2:
                return False

            sr = sel.ShapeRange
            orientation_restored = []

            # 4. Restore Orientation on Scratch Slide if dropped
            if self.preserve_orientation:
                # Restore Flips
                if orig_hflip and sr.HorizontalFlip == 0:
                    sr.Flip(MSO_FLIP_HORIZONTAL)
                    orientation_restored.append("H-Flip")

                if orig_vflip and sr.VerticalFlip == 0:
                    sr.Flip(MSO_FLIP_VERTICAL)
                    orientation_restored.append("V-Flip")

                # Restore Rotation
                if abs(sr.Rotation - orig_rot) > 0.1:
                    sr.Rotation = orig_rot
                    orientation_restored.append(f"Rot({orig_rot:.0f}°)")

            # 5. Cut converted shape and paste back into original container
            safe_clipboard_action(lambda: sr.Cut())
            time.sleep(self.delay)
            pasted_back = safe_clipboard_action(lambda: target_container.Shapes.Paste())
            time.sleep(self.delay)

            # 6. Re-apply any missing orientation/flips on the target container
            if self.preserve_orientation:
                if orig_hflip and pasted_back.HorizontalFlip == 0:
                    pasted_back.Flip(MSO_FLIP_HORIZONTAL)
                if orig_vflip and pasted_back.VerticalFlip == 0:
                    pasted_back.Flip(MSO_FLIP_VERTICAL)
                if abs(pasted_back.Rotation - orig_rot) > 0.1:
                    pasted_back.Rotation = orig_rot

            # 7. PIXEL-PERFECT POSITION ALIGNMENT
            # Measure pasted center and apply exact delta offset to eliminate clipboard shift
            pb_center_x, pb_center_y = self._get_center_coordinates(pasted_back)
            delta_x = orig_center_x - pb_center_x
            delta_y = orig_center_y - pb_center_y

            if abs(delta_x) > 0.01:
                pasted_back.IncrementLeft(delta_x)
            if abs(delta_y) > 0.01:
                pasted_back.IncrementTop(delta_y)

            # 8. Optional Auto-Ungrouping
            was_ungrouped = False
            if self.ungroup:
                try:
                    did_ungroup = False
                    for s_idx in range(pasted_back.Count, 0, -1):
                        s_item = pasted_back.Item(s_idx)
                        if s_item.Type == MSO_GROUP:
                            s_item.Ungroup()
                            did_ungroup = True
                    was_ungrouped = did_ungroup or (pasted_back.Count > 1)
                except Exception:
                    pass

            # 9. Delete original SVG
            shp.Delete()
            self.converted_count += 1

            self.conversion_log.append({
                "container": container_name,
                "shape": shape_name,
                "status": "Converted",
                "orientation": ", ".join(orientation_restored) if orientation_restored else "Preserved",
                "position": f"Exact (dx={delta_x:+.1f}, dy={delta_y:+.1f})",
                "ungrouped": was_ungrouped
            })
            return True

        except Exception as e:
            err_msg = str(e)
            self.errors.append((container_name, shape_name, err_msg))
            self.conversion_log.append({
                "container": container_name,
                "shape": shape_name,
                "status": f"Failed: {err_msg}",
                "orientation": "N/A",
                "position": "N/A",
                "ungrouped": False
            })
            return False

    def process_container(self, container, container_name):
        """
        Processes a shape container (SlideMaster, CustomLayout, or Slide),
        unrolling groups with SVGs and converting each SVG with orientation & position preservation.
        """
        try:
            shape_count = container.Shapes.Count
        except Exception:
            return

        # 1. Unroll any group containing SVGs
        for i in range(shape_count, 0, -1):
            try:
                shp = container.Shapes.Item(i)
                if shp.Type == MSO_GROUP:
                    has_svg = False
                    for g in range(1, shp.GroupItems.Count + 1):
                        if shp.GroupItems.Item(g).Type in MSO_GRAPHIC_TYPES:
                            has_svg = True
                            break
                    if has_svg:
                        shp.Ungroup()
                        time.sleep(self.delay)
            except Exception:
                continue

        # 2. Convert all graphic shapes
        try:
            shape_count = container.Shapes.Count
        except Exception:
            return

        for i in range(shape_count, 0, -1):
            try:
                shp = container.Shapes.Item(i)
                if shp.Type in MSO_GRAPHIC_TYPES:
                    self._convert_shape(shp, container, container_name)
            except Exception:
                continue

    def process_presentation(self, target_scope="all"):
        """
        Processes the presentation based on target scope.
        """
        # 1. Slide Masters & Custom Layouts
        if target_scope in ("all", "masters"):
            designs_count = self.pres.Designs.Count
            for d in range(1, designs_count + 1):
                design = self.pres.Designs.Item(d)
                master = design.SlideMaster
                master_name = f"SlideMaster #{d}"
                
                # Master root shapes
                self.process_container(master, f"{master_name} (Main)")

                # Custom Layout shapes
                layouts_count = master.CustomLayouts.Count
                for l in range(1, layouts_count + 1):
                    layout = master.CustomLayouts.Item(l)
                    layout_name = layout.Name or f"Layout #{l}"
                    self.process_container(layout, f"{master_name} -> {layout_name}")

        # 2. Regular Slides
        if target_scope in ("all", "slides"):
            slides_count = self.pres.Slides.Count
            # If scratch slide was created at the end, don't process it
            end_slide = (slides_count - 1) if self.scratch_slide else slides_count
            for s in range(1, end_slide + 1):
                slide = self.pres.Slides.Item(s)
                slide_name = f"Slide #{s}"
                self.process_container(slide, slide_name)

        # 3. Clean up scratch slide
        if self.scratch_slide:
            try:
                self.scratch_slide.Delete()
                self.scratch_slide = None
            except Exception:
                pass

def convert_presentation_svgs(input_path, output_path=None, target_scope="all", ungroup=False, preserve_orientation=True, make_backup=False, delay=0.05):
    """
    Main conversion routine.
    """
    in_file = Path(input_path).resolve()
    if not in_file.exists():
        console.print(f"[bold red]Error:[/bold red] Input file not found: {in_file}")
        return False

    if output_path:
        out_file = Path(output_path).resolve()
    else:
        out_file = in_file

    if make_backup and in_file == out_file:
        backup_file = in_file.with_suffix(in_file.suffix + ".bak")
        shutil.copy2(in_file, backup_file)
        console.print(f"[cyan]Backup created at:[/cyan] {backup_file}")

    ppt_app = get_powerpoint_app()
    console.print(f"[bold green]Opening presentation:[/bold green] {in_file.name}")
    
    try:
        presentation = ppt_app.Presentations.Open(str(in_file), 0, 0, 1)
    except Exception as e:
        console.print(f"[bold red]Error opening presentation:[/bold red] {e}")
        return False

    converter = SVGToShapeConverter(
        ppt_app=ppt_app,
        presentation=presentation,
        ungroup=ungroup,
        preserve_orientation=preserve_orientation,
        delay=delay
    )
    
    with console.status("[bold blue]Converting SVG graphics and aligning exact positions...[/bold blue]"):
        converter.process_presentation(target_scope=target_scope)

    # Save presentation
    try:
        if out_file != in_file:
            presentation.SaveAs(str(out_file))
            console.print(f"[bold green]Saved converted presentation to:[/bold green] {out_file}")
        else:
            presentation.Save()
            console.print(f"[bold green]Saved changes in-place to:[/bold green] {in_file}")
    except Exception as e:
        console.print(f"[bold red]Error saving presentation:[/bold red] {e}")
    finally:
        try:
            presentation.Close()
        except Exception:
            pass

    # Print summary table
    table = Table(title="VectoraKadavra Conversion Summary", header_style="bold magenta")
    table.add_column("Location / Container", style="cyan")
    table.add_column("Shape Name", style="white")
    table.add_column("Status", style="green")
    table.add_column("Orientation", style="magenta")
    table.add_column("Position", style="blue")
    table.add_column("Ungrouped", style="yellow")

    for log_entry in converter.conversion_log:
        status_style = "green" if "Converted" in log_entry["status"] else "red"
        table.add_row(
            log_entry["container"],
            log_entry["shape"],
            f"[{status_style}]{log_entry['status']}[/{status_style}]",
            log_entry["orientation"],
            log_entry["position"],
            "Yes" if log_entry["ungrouped"] else "No"
        )

    if converter.conversion_log:
        console.print(table)
    else:
        console.print(Panel("[yellow]No SVG graphics (msoGraphic / type 28/31) were found in the targeted scope.[/yellow]", title="Result"))

    console.print(f"\n[bold green]Total SVGs converted to shapes:[/bold green] {converter.converted_count}")
    if converter.errors:
        console.print(f"[bold red]Errors encountered:[/bold red] {len(converter.errors)}")

    return True

def main():
    parser = argparse.ArgumentParser(
        description="VectoraKadavra: Convert all SVGs in PowerPoint (including Slide Masters & Layouts) to native shapes with exact position and orientation preservation."
    )
    parser.add_argument("input", help="Path to input PowerPoint presentation (.pptx)")
    parser.add_argument("-o", "--output", help="Path to output presentation (default: overwrite input)")
    parser.add_argument(
        "-t", "--target",
        choices=["all", "masters", "slides"],
        default="all",
        help="Target scope: 'all' (default), 'masters' (Slide Master & Layouts only), or 'slides' (Slides only)"
    )
    parser.add_argument(
        "--ungroup",
        action="store_true",
        help="Automatically ungroup the converted shapes into individual vector paths"
    )
    parser.add_argument(
        "--no-preserve-orientation",
        dest="preserve_orientation",
        action="store_false",
        help="Disable automatic orientation/flip preservation"
    )
    parser.add_argument(
        "--backup",
        action="store_true",
        help="Create a .bak copy before modifying the file"
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=0.05,
        help="Delay in seconds between UI operations to ensure stability (default: 0.05)"
    )

    args = parser.parse_args()
    convert_presentation_svgs(
        input_path=args.input,
        output_path=args.output,
        target_scope=args.target,
        ungroup=args.ungroup,
        preserve_orientation=args.preserve_orientation,
        make_backup=args.backup,
        delay=args.delay
    )

if __name__ == "__main__":
    main()
