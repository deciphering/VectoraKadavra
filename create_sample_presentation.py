"""
Helper script to generate a test PowerPoint presentation containing SVGs
on the Slide Master, Custom Layout, and regular Slide for testing VectoraKadavra.
"""

import os
import sys
import win32com.client
from pathlib import Path

SAMPLE_SVG_CONTENT = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" width="100" height="100">
  <circle cx="50" cy="50" r="40" fill="#4F46E5" stroke="#312E81" stroke-width="4"/>
  <polygon points="50,20 60,40 82,42 66,58 70,80 50,68 30,80 34,58 18,42 40,40" fill="#FBBF24" />
</svg>
"""

SAMPLE_SVG_STAR = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" width="100" height="100">
  <rect x="10" y="10" width="80" height="80" rx="15" fill="#10B981" stroke="#047857" stroke-width="4"/>
  <path d="M30 50 L45 65 L70 35" fill="none" stroke="#FFFFFF" stroke-width="8" stroke-linecap="round" stroke-linejoin="round"/>
</svg>
"""

def create_sample_pptx(output_path="test_presentation.pptx"):
    out_file = Path(output_path).resolve()
    temp_svg_1 = Path("temp_icon1.svg").resolve()
    temp_svg_2 = Path("temp_icon2.svg").resolve()

    temp_svg_1.write_text(SAMPLE_SVG_CONTENT, encoding="utf-8")
    temp_svg_2.write_text(SAMPLE_SVG_STAR, encoding="utf-8")

    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True

    pres = ppt_app.Presentations.Add()

    try:
        # 1. Add SVG to Slide Master
        ppt_app.ActiveWindow.ViewType = 2  # ppViewSlideMaster
        master = pres.Designs.Item(1).SlideMaster
        master.Shapes.AddPicture(str(temp_svg_1), LinkToFile=0, SaveWithDocument=-1, Left=50, Top=50, Width=80, Height=80)

        # 2. Add SVG to Custom Layout #1
        layout = master.CustomLayouts.Item(1)
        layout.Shapes.AddPicture(str(temp_svg_2), LinkToFile=0, SaveWithDocument=-1, Left=200, Top=50, Width=80, Height=80)

        # 3. Add SVG to Slide 1
        ppt_app.ActiveWindow.ViewType = 9  # ppViewNormal
        slide_layout = pres.SlideMaster.CustomLayouts.Item(1)
        slide = pres.Slides.AddSlide(1, slide_layout)
        slide.Shapes.AddPicture(str(temp_svg_1), LinkToFile=0, SaveWithDocument=-1, Left=350, Top=200, Width=100, Height=100)

        pres.SaveAs(str(out_file))
        print(f"Sample presentation created at: {out_file}")
    finally:
        pres.Close()
        # Clean up temporary SVG files
        if temp_svg_1.exists():
            temp_svg_1.unlink()
        if temp_svg_2.exists():
            temp_svg_2.unlink()

if __name__ == "__main__":
    create_sample_pptx()
