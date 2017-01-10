import sys, win32com.client, MSO, MSPPT, time, random
g = globals()
for c in dir(MSO.constants):    g[c] = getattr(MSO.constants, c)
for c in dir(MSPPT.constants):  g[c] = getattr(MSPPT.constants, c)

def RGBtoInt(r, g, b):
    return r + g*256 + b*256*256

#--------- Definitions
group1_x_start = 50;    group1_x_stop=200
group1_y_start = 50;    group1_y_stop=200

group2_x_start = 300;   group2_x_stop=500
group2_y_start = 50;    group2_y_stop=200

group3_x_start = 100;   group3_x_stop=350
group3_y_start = 300;    group3_y_stop=450

yellow_color = RGBtoInt(255, 255, 0)
blue_color = RGBtoInt(0, 0, 255)
red_color = RGBtoInt(255, 0, 0)

#----- Slide Generation

# Open PowerPoint
Application = win32com.client.Dispatch("PowerPoint.Application")

# Add a presentation
Presentation = Application.Presentations.Add()

# Add a slide with a blank layout (12 stands for blank layout)
Base = Presentation.Slides.Add(1, 12)

samples = []
total_samples = 12

sample_ranges = []
sample_ranges.append(((group1_x_start, group1_x_stop),(group1_y_start, group1_y_stop)))
sample_ranges.append(((group2_x_start, group2_x_stop),(group2_y_start, group2_y_stop)))
sample_ranges.append(((group3_x_start, group3_x_stop),(group3_y_start, group3_y_stop)))



for i in range(total_samples):
    group_index = i/(total_samples/3)
    rand_x_coord = random.randrange(sample_ranges[group_index][0][0],sample_ranges[group_index][0][1])
    rand_y_coord = random.randrange(sample_ranges[group_index][1][0],sample_ranges[group_index][1][1])
    samples.append(Base.Shapes.AddShape(9, rand_x_coord, rand_y_coord, 40, 40))
    samples[i].Fill.ForeColor.RGB = RGBtoInt(2*i, 2*i, 2*i)
    samples[i].Line.ForeColor.RGB = RGBtoInt(0,0,0)
    samples[i].Line.Weight = 1
    samples[i].TextFrame.TextRange.Text = str(int(samples[i].Top))
    samples[i].TextFrame.TextRange.font.size = 12

line = Base.Shapes.AddLine(275, 50, 275, 500)
line.line.foreColor.RGB = 0
line.line.weight = 3.5
line.line.EndArrowheadStyle  = MSO.constants.msoArrowheadTriangle

line2 = Base.Shapes.AddLine(30, 275, 550, 275)
line2.line.foreColor.RGB = 0
line2.line.weight = 3.5
line2.line.EndArrowheadStyle  = MSO.constants.msoArrowheadTriangle

    # Add an oval. Shape 9 is an oval.
#oval = Base.Shapes.AddShape(9, 2, 2, 2, 2)
#oval = Base.Shapes.AddShape(9, 0, 100, 100, 100)
#line = Base.Shapes.AddShape(0x6d, 0, 100, 100, 100)
#oval2 = Base.Shapes.AddShape(9, 100, 0, 50, 50)
#oval.Fill.ForeColor.RGB = RGB(255, 0, 0)
#oval.top = 200


#for i in range(1, 50):
#    Base.Shapes.AddShape(i, 100, i*5, 5, 5)
#line = Base.Shapes.AddLine(10, 300, 200, 300)
#line.line.foreColor.RGB = 0
#line.line.dashstyle = MSO.constants.msoLineThickThin

#oval = Base.Shapes.AddShape(9, 0, 100, 100, 100)
#oval.Fill.ForeColor.RGB = RGB(255, 0, 0)