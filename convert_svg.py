"""
VectoraKadavra: Convert SVG graphics to native PowerPoint shapes programmatically.
Focuses on Slide Masters, Custom Layouts, and regular Slides.
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

# PowerPoint Constants
MSO_GRAPHIC_TYPES = (28, 31)  # SVG / Office Graphic shape types across Office versions
MSO_GROUP = 6                 # Group shape
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

class SVGToShapeConverter:
    def __init__(self, ppt_app, ungroup=False, delay=0.05):
        self.ppt_app = ppt_app
        self.ungroup = ungroup
        self.delay = delay
        self.converted_count = 0
        self.errors = []
        self.conversion_log = []

    def _convert_shape(self, shp, container_name):
        """Attempts to select the shape and trigger the 'SVGEdit' (Convert to Shape) command."""
        try:
            shape_name = shp.Name
        except Exception:
            shape_name = "Unknown Graphic"

        try:
            # Select the target graphic shape in the active view
            shp.Select()
            if self.delay > 0:
                time.sleep(self.delay)
            
            # Execute native PowerPoint "Convert to Shape"
            self.ppt_app.CommandBars.ExecuteMso("SVGEdit")
            if self.delay > 0:
                time.sleep(self.delay)
            
            self.converted_count += 1
            log_item = {
                "container": container_name,
                "shape": shape_name,
                "status": "Converted",
                "ungrouped": False
            }
            self.conversion_log.append(log_item)

            # Optional automatic ungrouping
            if self.ungroup:
                try:
                    sel = self.ppt_app.ActiveWindow.Selection
                    if sel.Type == 2:  # ppSelectionShapes
                        did_ungroup = False
                        # If it created multiple shapes or nested groups
                        for s_idx in range(sel.ShapeRange.Count, 0, -1):
                            s_item = sel.ShapeRange.Item(s_idx)
                            if s_item.Type == MSO_GROUP:
                                s_item.Ungroup()
                                did_ungroup = True
                        log_item["ungrouped"] = did_ungroup or (sel.ShapeRange.Count > 1)
                except Exception:
                    pass

            return True
        except Exception as e:
            err_msg = str(e)
            self.errors.append((container_name, shape_name, err_msg))
            self.conversion_log.append({
                "container": container_name,
                "shape": shape_name,
                "status": f"Failed: {err_msg}",
                "ungrouped": False
            })
            return False

    def process_container(self, container, container_name):
        """
        Iterates backwards through shapes in a container (SlideMaster, CustomLayout, or Slide)
        and converts any msoGraphic (SVG) shapes to native shapes.
        """
        # Ensure the container itself is selected/active in the current view
        try:
            if hasattr(container, "Select"):
                container.Select()
                if self.delay > 0:
                    time.sleep(self.delay)
        except Exception:
            pass

        try:
            shape_count = container.Shapes.Count
        except Exception:
            return

        # Iterate in reverse because converting shapes replaces the graphic in the collection
        for i in range(shape_count, 0, -1):
            try:
                shp = container.Shapes.Item(i)
                if shp.Type in MSO_GRAPHIC_TYPES:
                    self._convert_shape(shp, container_name)
            except Exception:
                continue

    def process_presentation(self, presentation, target_scope="all"):
        """
        Processes the presentation based on the chosen scope:
        - 'masters': Only Slide Masters and Custom Layouts
        - 'slides': Only normal Slides
        - 'all': Slide Masters, Custom Layouts, and Slides
        """
        # 1. Process Slide Masters and Custom Layouts
        if target_scope in ("all", "masters"):
            try:
                self.ppt_app.ActiveWindow.ViewType = PP_VIEW_SLIDE_MASTER
                if self.delay > 0:
                    time.sleep(0.1)
            except Exception:
                pass

            designs_count = presentation.Designs.Count
            for d in range(1, designs_count + 1):
                design = presentation.Designs.Item(d)
                master = design.SlideMaster
                master_name = f"SlideMaster #{d}"
                
                # Master-level shapes
                self.process_container(master, f"{master_name} (Main)")

                # Custom Layout shapes
                layouts_count = master.CustomLayouts.Count
                for l in range(1, layouts_count + 1):
                    layout = master.CustomLayouts.Item(l)
                    layout_name = layout.Name or f"Layout #{l}"
                    self.process_container(layout, f"{master_name} -> {layout_name}")

        # 2. Process Regular Slides
        if target_scope in ("all", "slides"):
            try:
                self.ppt_app.ActiveWindow.ViewType = PP_VIEW_NORMAL
                if self.delay > 0:
                    time.sleep(0.1)
            except Exception:
                pass

            slides_count = presentation.Slides.Count
            for s in range(1, slides_count + 1):
                slide = presentation.Slides.Item(s)
                slide_name = f"Slide #{s}"
                self.process_container(slide, slide_name)

def convert_presentation_svgs(input_path, output_path=None, target_scope="all", ungroup=False, make_backup=False, delay=0.05):
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
        # Open presentation in PowerPoint (WithWindow=msoTrue is required for UI commands)
        presentation = ppt_app.Presentations.Open(str(in_file), 0, 0, 1)
    except Exception as e:
        console.print(f"[bold red]Error opening presentation:[/bold red] {e}")
        return False

    converter = SVGToShapeConverter(ppt_app, ungroup=ungroup, delay=delay)
    
    with console.status("[bold blue]Scanning and converting SVG graphics to native shapes...[/bold blue]"):
        converter.process_presentation(presentation, target_scope=target_scope)

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
    table.add_column("Ungrouped", style="yellow")

    for log_entry in converter.conversion_log:
        status_style = "green" if "Converted" in log_entry["status"] else "red"
        table.add_row(
            log_entry["container"],
            log_entry["shape"],
            f"[{status_style}]{log_entry['status']}[/{status_style}]",
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
        description="VectoraKadavra: Convert all SVGs in PowerPoint (including Slide Masters & Layouts) to native shapes."
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
        make_backup=args.backup,
        delay=args.delay
    )

if __name__ == "__main__":
    main()
