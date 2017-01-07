import sys, win32com.client, MSO, MSPPT, time
g = globals()
for c in dir(MSO.constants):    g[c] = getattr(MSO.constants, c)
for c in dir(MSPPT.constants):  g[c] = getattr(MSPPT.constants, c)

def RGB(r, g, b):
    return r + g*256 + b*256*256

# Open PowerPoint
Application = win32com.client.Dispatch("PowerPoint.Application")

# Add a presentation
Presentation = Application.Presentations.Add()

# Add a slide with a blank layout (12 stands for blank layout)
Base = Presentation.Slides.Add(1, 12)

# Add an oval. Shape 9 is an oval.
#oval = Base.Shapes.AddShape(9, 2, 2, 2, 2)
#oval = Base.Shapes.AddShape(9, 0, 100, 100, 100)
#line = Base.Shapes.AddShape(0x6d, 0, 100, 100, 100)
#oval2 = Base.Shapes.AddShape(9, 100, 0, 50, 50)
#oval.Fill.ForeColor.RGB = RGB(255, 0, 0)
#oval.top = 200

#for i in range(1, 50):
#    Base.Shapes.AddShape(i, 100, i*5, 5, 5)
line = Base.Shapes.AddLine(10, 300, 200, 300)
line.line.foreColor.RGB = 0
line.line.dashstyle = MSO.constants.msoLineThickThin

Base = Presentation.Slides.Add(2, 1)
oval = Base.Shapes.AddShape(9, 0, 100, 100, 100)
oval.Fill.ForeColor.RGB = RGB(255, 0, 0)