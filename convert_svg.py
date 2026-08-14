"""
VectoraKadavra: Convert SVG graphics to native PowerPoint shapes programmatically.
Strict workflow:
1. Capture exact XY position, dimensions, orientation (rotation & flips), and crop metadata before conversion.
2. Use duplication to extract the exact true uncropped physical dimensions and position.
3. Convert to shape and position at full 1:1 true scale to prevent squishing and stretching.
4. If cropped, cull paths outside the unrotated visible crop window cleanly.
5. Apply original rotation and final cleanup.

Supports Headless/Minimized mode so PowerPoint runs seamlessly in the background.
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
MSO_GRAPHIC_TYPES = (28, 31)  # SVG / Graphic shape types
MSO_GROUP = 6                 # Group shape
MSO_SHAPE_RECTANGLE = 1       # msoShapeRectangle
MSO_FLIP_HORIZONTAL = 0       # msoFlipHorizontal
MSO_FLIP_VERTICAL = 1         # msoFlipVertical
PP_VIEW_NORMAL = 9            # Normal View

def hide_powerpoint_window(ppt_app):
    """Completely hides the PowerPoint window using off-screen positioning and taskbar removal."""
    try:
        import win32gui
        import win32con
        
        # Set a unique caption so we only hide OUR instance, not the user's other PPT windows
        ppt_app.Caption = "VectoraKadavra_Processing"
        hwnd = win32gui.FindWindow("PPTFrameClass", "VectoraKadavra_Processing")
        
        if hwnd:
            # Move deep off-screen without deactivating the view
            win32gui.SetWindowPos(
                hwnd,
                win32con.HWND_BOTTOM,
                -20000,
                -20000,
                200,
                200,
                win32con.SWP_NOACTIVATE
            )
            
            # Remove from taskbar and Alt-Tab by swapping APPWINDOW for TOOLWINDOW
            ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            ex_style = ex_style & ~win32con.WS_EX_APPWINDOW
            ex_style = ex_style | win32con.WS_EX_TOOLWINDOW
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style)
    except Exception:
        pass

def get_powerpoint_app(headless=True):
    """Initializes and returns the PowerPoint COM Application instance."""
    try:
        import win32com.client
    except ImportError:
        console.print("[bold red]Error:[/bold red] pywin32 is not installed. Run 'pip install pywin32' first.")
        sys.exit(1)

    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        if headless:
            hide_powerpoint_window(ppt_app)
        else:
            ppt_app.Visible = True
        return ppt_app
    except Exception as e:
        console.print(f"[bold red]Failed to launch Microsoft PowerPoint COM Automation:[/bold red] {e}")
        console.print("[yellow]Make sure Microsoft PowerPoint (Office 365 or 2019+) is installed on this machine.[/yellow]")
        sys.exit(1)

def safe_com_action(action_func, max_retries=10, delay=0.1):
    """Executes a COM action with retries to handle Windows clipboard/COM locks."""
    for attempt in range(max_retries):
        try:
            return action_func()
        except Exception as e:
            if attempt == max_retries - 1:
                raise e
            time.sleep(delay * (attempt + 1))

class SVGToShapeConverter:
    def __init__(self, ppt_app, presentation, ungroup=False, headless=True, delay=0.05):
        self.ppt_app = ppt_app
        self.pres = presentation
        self.ungroup = ungroup
        self.headless = headless
        self.delay = delay
        self.converted_count = 0
        self.errors = []
        self.conversion_log = []
        self.scratch_slide = None

    def _get_or_create_scratch_slide(self):
        """Creates a dedicated scratch slide in Normal view for reliable SVGEdit execution."""
        if self.scratch_slide is None:
            self.ppt_app.ActiveWindow.ViewType = PP_VIEW_NORMAL
            if self.headless:
                hide_powerpoint_window(self.ppt_app)
            time.sleep(0.1)
            blank_layout = self.pres.SlideMaster.CustomLayouts.Item(7)
            self.scratch_slide = self.pres.Slides.AddSlide(self.pres.Slides.Count + 1, blank_layout)
        return self.scratch_slide

    def _apply_vector_crop(self, container, target_shape, orig_left, orig_top, orig_width, orig_height):
        """
        Safely culls subshapes that are entirely outside the visible crop box.
        This runs while target_shape is at 0 rotation, ensuring bounding boxes are perfectly orthogonal.
        """
        try:
            tol = 1.0
            vis_min_x = orig_left - tol
            vis_max_x = orig_left + orig_width + tol
            vis_min_y = orig_top - tol
            vis_max_y = orig_top + orig_height + tol

            if target_shape.Type == MSO_GROUP:
                ungrouped = target_shape.Ungroup()
                time.sleep(self.delay)
                items = [ungrouped.Item(i) for i in range(1, ungrouped.Count + 1)]
                kept_names = []

                for s in items:
                    try:
                        s_left = s.Left
                        s_top = s.Top
                        s_right = s_left + s.Width
                        s_bottom = s_top + s.Height

                        # Completely outside visible unrotated crop window -> delete safely
                        if s_right <= vis_min_x or s_left >= vis_max_x or s_bottom <= vis_min_y or s_top >= vis_max_y:
                            s.Delete()
                        else:
                            kept_names.append(s.Name)
                    except Exception:
                        pass

                if kept_names:
                    return container.Shapes.Range(kept_names).Group()
                return target_shape
            else:
                return target_shape
        except Exception:
            return target_shape

    def _convert_single_svg(self, shp, target_container, container_name):
        """
        Converts an SVG shape to native shapes following the strict workflow.
        """
        try:
            shape_name = shp.Name
        except Exception:
            shape_name = "SVG Graphic"

        try:
            # =========================================================================
            # STEP 1: Capture exact XY position, dimensions, & orientation BEFORE conversion
            # =========================================================================
            orig_left = float(shp.Left)
            orig_top = float(shp.Top)
            orig_width = float(shp.Width)
            orig_height = float(shp.Height)
            orig_rot = float(getattr(shp, 'Rotation', 0.0))
            orig_hflip = (getattr(shp, 'HorizontalFlip', 0) == -1)
            orig_vflip = (getattr(shp, 'VerticalFlip', 0) == -1)

            # =========================================================================
            # STEP 2: Use Duplicate to extract true uncropped physical bounds
            # =========================================================================
            crop_info = []
            has_crop = False
            full_left, full_top, full_width, full_height = orig_left, orig_top, orig_width, orig_height

            try:
                pf = shp.PictureFormat
                cl = float(getattr(pf, 'CropLeft', 0.0))
                ct = float(getattr(pf, 'CropTop', 0.0))
                cr = float(getattr(pf, 'CropRight', 0.0))
                cb = float(getattr(pf, 'CropBottom', 0.0))
                
                if abs(cl) > 0.01 or abs(ct) > 0.01 or abs(cr) > 0.01 or abs(cb) > 0.01:
                    has_crop = True
                    crop_info.append(f"L:{cl:.0f} T:{ct:.0f} R:{cr:.0f} B:{cb:.0f}")

                    # Duplicate and uncrop to let PowerPoint calculate the exact physical full size
                    dup = shp.Duplicate()
                    if hasattr(dup, "Count") and dup.Count >= 1:
                        dup = dup.Item(1)
                    
                    # Snap duplicate back to original position to eliminate PowerPoint's default +12px offset
                    dup.Left = orig_left
                    dup.Top = orig_top

                    dup.PictureFormat.CropLeft = 0.0
                    dup.PictureFormat.CropTop = 0.0
                    dup.PictureFormat.CropRight = 0.0
                    dup.PictureFormat.CropBottom = 0.0
                    
                    full_left = float(dup.Left)
                    full_top = float(dup.Top)
                    full_width = float(dup.Width)
                    full_height = float(dup.Height)
                    
                    try:
                        dup.Delete()
                    except Exception:
                        pass
            except Exception:
                pass

            # =========================================================================
            # STEP 3: Convert SVG to vector
            # =========================================================================
            scratch = self._get_or_create_scratch_slide()
            safe_com_action(lambda: shp.Copy())
            scratch.Select()
            time.sleep(self.delay)

            pasted_range = safe_com_action(lambda: scratch.Shapes.Paste())
            pasted_shp = pasted_range.Item(1)

            pasted_shp.Select()
            time.sleep(self.delay)
            self.ppt_app.CommandBars.ExecuteMso("SVGEdit")
            time.sleep(self.delay)

            sel = self.ppt_app.ActiveWindow.Selection
            if sel.Type != 2:
                return False
            sr = sel.ShapeRange

            safe_com_action(lambda: sr.Cut())
            time.sleep(self.delay)
            pasted_back = safe_com_action(lambda: target_container.Shapes.Paste())
            time.sleep(self.delay)

            shp.Delete()

            if pasted_back.Count > 1:
                target_shape = pasted_back.Group()
            else:
                target_shape = pasted_back.Item(1)

            # =========================================================================
            # STEP 4: Position at TRUE aspect ratio (no squishing)
            # =========================================================================
            orientation_actions = []

            if orig_hflip and target_shape.HorizontalFlip == 0:
                target_shape.Flip(MSO_FLIP_HORIZONTAL)
                orientation_actions.append("H-Flip")
            if orig_vflip and target_shape.VerticalFlip == 0:
                target_shape.Flip(MSO_FLIP_VERTICAL)
                orientation_actions.append("V-Flip")

            # Force 0 rotation during sizing and culling for orthogonal accuracy
            target_shape.Rotation = 0.0

            # Set true uncropped physical dimensions
            target_shape.Left = full_left
            target_shape.Top = full_top
            target_shape.Width = full_width
            target_shape.Height = full_height

            # =========================================================================
            # STEP 5: Apply safe culling ONLY if cropped
            # =========================================================================
            if has_crop:
                target_shape = self._apply_vector_crop(
                    target_container,
                    target_shape,
                    orig_left,
                    orig_top,
                    orig_width,
                    orig_height
                )

            # =========================================================================
            # STEP 6: Apply final rotation and cleanup
            # =========================================================================
            if abs(target_shape.Rotation - orig_rot) > 0.1:
                target_shape.Rotation = orig_rot
                orientation_actions.append(f"Rot({orig_rot:.0f}°)")

            was_ungrouped = False
            if self.ungroup:
                try:
                    if target_shape.Type == MSO_GROUP:
                        target_shape.Ungroup()
                        was_ungrouped = True
                except Exception:
                    pass

            self.converted_count += 1
            self.conversion_log.append({
                "container": container_name,
                "shape": shape_name,
                "status": "Converted",
                "orientation": ", ".join(orientation_actions) if orientation_actions else "Preserved",
                "crop": ", ".join(crop_info) if crop_info else "None",
                "xy": f"({target_shape.Left:.1f}, {target_shape.Top:.1f})",
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
                "crop": "N/A",
                "xy": "N/A",
                "ungrouped": False
            })
            return False

    def _group_contains_svg(self, grp_shape):
        try:
            for g in range(1, grp_shape.GroupItems.Count + 1):
                item = grp_shape.GroupItems.Item(g)
                if item.Type in MSO_GRAPHIC_TYPES:
                    return True
                if item.Type == MSO_GROUP:
                    if self._group_contains_svg(item):
                        return True
        except Exception:
            pass
        return False

    def process_container(self, container, container_name):
        """
        Processes all shapes in a container (SlideMaster, CustomLayout, or Slide).
        """
        # 1. Recursively unpack any groups that contain SVGs so we don't miss them
        # and so they inherit absolute physical transforms before conversion.
        while True:
            ungrouped_any = False
            try:
                shape_count = container.Shapes.Count
            except Exception:
                break
                
            for i in range(shape_count, 0, -1):
                try:
                    shp = container.Shapes.Item(i)
                    if shp.Type == MSO_GROUP:
                        if self._group_contains_svg(shp):
                            shp.Ungroup()
                            time.sleep(self.delay)
                            ungrouped_any = True
                            break
                except Exception:
                    continue
                    
            if not ungrouped_any:
                break

        # 2. Convert all graphic shapes
        try:
            shape_count = container.Shapes.Count
        except Exception:
            return

        for i in range(shape_count, 0, -1):
            try:
                shp = container.Shapes.Item(i)
                if shp.Type in MSO_GRAPHIC_TYPES:
                    self._convert_single_svg(shp, container, container_name)
            except Exception:
                continue

    def process_presentation(self, target_scope="all"):
        if target_scope in ("all", "masters"):
            designs_count = safe_com_action(lambda: self.pres.Designs.Count)
            for d in range(1, designs_count + 1):
                design = safe_com_action(lambda: self.pres.Designs.Item(d))
                master = safe_com_action(lambda: design.SlideMaster)
                master_name = f"SlideMaster #{d}"
                self.process_container(master, f"{master_name} (Main)")

                layouts_count = safe_com_action(lambda: master.CustomLayouts.Count)
                for l in range(1, layouts_count + 1):
                    layout = safe_com_action(lambda: master.CustomLayouts.Item(l))
                    try:
                        layout_name = safe_com_action(lambda: layout.Name)
                    except Exception:
                        layout_name = None
                    layout_name = layout_name or f"Layout #{l}"
                    self.process_container(layout, f"{master_name} -> {layout_name}")

        if target_scope in ("all", "slides"):
            slides_count = safe_com_action(lambda: self.pres.Slides.Count)
            end_slide = (slides_count - 1) if self.scratch_slide else slides_count
            for s in range(1, end_slide + 1):
                slide = safe_com_action(lambda: self.pres.Slides.Item(s))
                slide_name = f"Slide #{s}"
                self.process_container(slide, slide_name)

        if self.scratch_slide:
            try:
                self.scratch_slide.Delete()
                self.scratch_slide = None
            except Exception:
                pass

def convert_presentation_svgs(input_path, output_path=None, target_scope="all", ungroup=False, make_backup=False, headless=True, delay=0.05):
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

    ppt_app = get_powerpoint_app(headless=headless)
    console.print(f"[bold green]Processing presentation in {'headless' if headless else 'visible'} mode:[/bold green] {in_file.name}")
    
    try:
        presentation = ppt_app.Presentations.Open(str(in_file), 0, 0, 1)
        if headless:
            hide_powerpoint_window(ppt_app)
    except Exception as e:
        console.print(f"[bold red]Error opening presentation:[/bold red] {e}")
        return False

    converter = SVGToShapeConverter(
        ppt_app=ppt_app,
        presentation=presentation,
        ungroup=ungroup,
        headless=headless,
        delay=delay
    )
    
    with console.status("[bold blue]Converting SVG graphics and extracting true uncropped physics...[/bold blue]"):
        converter.process_presentation(target_scope=target_scope)

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
        if headless:
            try:
                ppt_app.Quit()
            except Exception:
                pass

    table = Table(title="VectoraKadavra Conversion Summary", header_style="bold magenta")
    table.add_column("Location / Container", style="cyan")
    table.add_column("Shape Name", style="white")
    table.add_column("Status", style="green")
    table.add_column("Orientation", style="magenta")
    table.add_column("Crop", style="yellow")
    table.add_column("Final (X, Y)", style="blue")
    table.add_column("Ungrouped", style="dim")

    for log_entry in converter.conversion_log:
        status_style = "green" if "Converted" in log_entry["status"] else "red"
        table.add_row(
            log_entry["container"],
            log_entry["shape"],
            f"[{status_style}]{log_entry['status']}[/{status_style}]",
            log_entry["orientation"],
            log_entry.get("crop", "None"),
            log_entry.get("xy", "N/A"),
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
        description="VectoraKadavra: Convert all SVGs in PowerPoint (including Slide Masters & Layouts) to native shapes headlessly with true physical aspect ratio preservation."
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
        help="Automatically ungroup the converted shapes into individual vector paths after positioning"
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Show PowerPoint window during execution (by default, runs in headless/hidden mode)"
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
        headless=not args.visible,
        delay=args.delay
    )

if __name__ == "__main__":
    main()
