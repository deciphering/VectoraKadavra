import win32com.client

ppt = win32com.client.Dispatch('PowerPoint.Application')
pres = ppt.Presentations.Open(r'C:\Antigravity Projects\VectoraKadavra\test_presentation.pptx')
print(f"Designs: {pres.Designs.Count}")
for d in range(1, pres.Designs.Count + 1):
    m = pres.Designs.Item(d).SlideMaster
    print(f"Master {d} shapes: {m.Shapes.Count}")
    for s in range(1, m.Shapes.Count + 1):
        shp = m.Shapes.Item(s)
        print(f"  Master Shape {s}: name={shp.Name}, type={shp.Type}")
    for l in range(1, m.CustomLayouts.Count + 1):
        layout = m.CustomLayouts.Item(l)
        for s in range(1, layout.Shapes.Count + 1):
            shp = layout.Shapes.Item(s)
            print(f"  Layout {l} ({layout.Name}) Shape {s}: name={shp.Name}, type={shp.Type}")
for s_idx in range(1, pres.Slides.Count + 1):
    slide = pres.Slides.Item(s_idx)
    print(f"Slide {s_idx} shapes: {slide.Shapes.Count}")
    for s in range(1, slide.Shapes.Count + 1):
        shp = slide.Shapes.Item(s)
        print(f"  Slide Shape {s}: name={shp.Name}, type={shp.Type}")
pres.Close()
