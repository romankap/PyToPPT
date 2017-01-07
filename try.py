import win32com.client

# Open PowerPoint
Application = win32com.client.Dispatch("PowerPoint.Application")

# Add a presentation
Presentation = Application.Presentations.Add()

# Add a slide with a blank layout (12 stands for blank layout)
Base = Presentation.Slides.Add(1, 12)

# Add an oval. Shape 9 is an oval.
oval = Base.Shapes.AddShape(9, 2, 2, 2, 2)
oval = Base.Shapes.AddShape(9, 300, 200, 100, 100)