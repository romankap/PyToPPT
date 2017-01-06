#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Shengdong Zhao
#
# Created:     22/07/2012
# Copyright:   (c) Shengdong Zhao 2012
# Licence:     <your licence>
#-------------------------------------------------------------------------------

#import all the modules. MSPPTUtil is a module that allows you to create a
# presentation and insert title slide, outline slide, and slide with pictures
import sys, win32com.client, MSO, MSPPT, MSPPTUtil
g = globals()
for c in dir(MSO.constants):    g[c] = getattr(MSO.constants, c)
for c in dir(MSPPT.constants):  g[c] = getattr(MSPPT.constants, c)

# Step 1: create a set of slides
# initialize a list of slides
slides = []
# initialize the different type of slide object
# In the text below, it shows how to create a cover slide, an outline slide, a text+picture slide, a picture slide, and a simple table slide
slide = MSPPTUtil.PptCover('this is a subtitle', 'title')
slide1 = MSPPTUtil.PptOutline('this is main point\n\tbullet point 1\n\tbullet point 2', 'title')
slide2 = MSPPTUtil.PptTextPicture('text', 'C:\\Users\\Shengdong Zhao\\Downloads\\treemap\\photo.JPG', 'title with photo')
slide3 = MSPPTUtil.PptPicture('C:\\Users\\Shengdong Zhao\\Downloads\\treemap\\photo.JPG', 'title with photo')
slide4 = MSPPTUtil.PptBasicTable([[1,2],[2,3]], ['column1','column2'],['row1','row2'],'slide title')

# add slide(s) to the slides list
slides.append(slide)
slides.append(slide1)
slides.append(slide2)
slides.append(slide3)
slides.append(slide4)

# step 2: create a presentation and add the slides to the presentation. Set the slider master's footer to be a string
# set the datetime field to be true, and set the gradient of the slide
presentation = MSPPTUtil.PptPresentation(slides, footer='shen', date_time=True, preset_gradient=(6, 1, 10))

# step 3: show the slide
presentation.create_show()

# Now you can save the slide somewhere
presentation.save_as('C:\\Dev\\PyToPPT\\roman_try.ppt')

# Close the presentation
presentation.close()




