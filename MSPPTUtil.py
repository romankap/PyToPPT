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

import sys, win32com.client, MSO, MSPPT, time
g = globals()
for c in dir(MSO.constants):    g[c] = getattr(MSO.constants, c)
for c in dir(MSPPT.constants):  g[c] = getattr(MSPPT.constants, c)

class PptPresentation(object):
	'''Simple Python class for creating PowerPoint presentations.
	'''
	def __init__(self, slide_list, **kwargs):
		self.slide_list = slide_list
		self.kwargs = kwargs
		self.ppt = None
		self.ppt_presentation = None
	def create_show(self, visible = True):
		import win32com.client as win32
		ppt = win32.gencache.EnsureDispatch('Powerpoint.Application')
		self.ppt = ppt
		ppt.Visible = visible  #bool
		presentation = ppt.Presentations.Add()
		self.presentation = presentation
		#set a few options
		kwargs = self.kwargs
		slide_master = presentation.SlideMaster
		if 'footer' in kwargs: #string
			slide_master.HeadersFooters.Footer.Text = kwargs['footer']
		if 'date_time' in kwargs: #bool
			#turn on or off the DateAndTime field
			slide_master.HeadersFooters.DateAndTime.UseFormat = kwargs['date_time']
		if 'preset_gradient' in kwargs: #3-tuple, e.g. (6, 1, 10)
			slide_master.Background.Fill.PresetGradient(*kwargs['preset_gradient'])
		#insert last slide first
		for slide in reversed(self.slide_list):
			slide.add_slide(presentation)
			if visible:
				slide.select() #bring slide to front
			slide.format_slide()
	def save_as(self, file_name):
		self.presentation.SaveAs(file_name)
	def close(self):
		self.presentation.Close()
	def ppt_quit(self):
		self.ppt.Quit()



class PptTitleSlide(object):
	def add_slide(self, presentation, slide_num=1):
		#create slide
		slide = presentation.Slides.Add(slide_num, self.layout)
		self.slide = slide
		time.sleep(0.1)
	def select(self):
		self.slide.Select()
	def format_slide(self):
		self.format_content()
		self.format_title()
	def format_title(self):
		slide = self.slide
		if self.title:
			#add title
			#title_range = slide.Shapes[0].TextFrame.TextRange
			title_range = slide.Shapes.Title.TextFrame.TextRange
			title_range.Text = self.title
			title_range.Font.Bold = True
			time.sleep(0.10)
		else:
			#slide.Shapes[0].Delete()
			slide.Shapes.Title.Delete()
			slide.Shapes[0].Top = 20
			slide.Shapes[0].Height = 13*36

def insert_outline(text_range, text):
	for line in text.split("\n"):
		line = line.rstrip()
		indent = 1 + line.count("\t")
		line = line.lstrip()
		if line:
			line = text_range.InsertAfter(line+"\r\n")
			line.IndentLevel = indent
			time.sleep(0.10)


class PptCover(PptTitleSlide):
	'''Bulleted outline with optional title (above).
	Each line of text is bulleted.
	Level of indentation (real tabs) determines outline level.
	'''
	def __init__(self, text, title=''):
		self.subtitle = text
		self.title = title
		self.slide = None
		self.layout = ppLayoutTitle
	def format_content(self):
		slide = self.slide
		#add text
		text = self.subtitle
		text_range = slide.Shapes[1].TextFrame.TextRange
		#text_range.InsertAfter(self.text)
		insert_outline(text_range, text)



class PptOutline(PptTitleSlide):
	'''Bulleted outline with optional title (above).
	Each line of text is bulleted.
	Level of indentation (real tabs) determines outline level.
	'''
	def __init__(self, text, title=''):
		self.text = text
		self.title = title
		self.slide = None
		self.layout = ppLayoutText
	def format_content(self):
		slide = self.slide
		#add text
		text = self.text
		text_range = slide.Shapes[1].TextFrame.TextRange
		#text_range.InsertAfter(self.text)
		insert_outline(text_range, text)

class PptPicture(PptTitleSlide):
	'''Slide with picture inserted from file and title.

	:see: http://msdn2.microsoft.com/en-us/library/aa211638(office.11).aspx
	'''
	def __init__(self, file_path, title):
		self.file_path = file_path
		self.title = title
		self.slide = None
		self.layout = ppLayoutTitleOnly
	def format_content(self):
		slide = self.slide
		#title_range = slide.Shapes[0].TextFrame.TextRange
		title_range = slide.Shapes.Title.TextFrame.TextRange
		title_range.Text = self.title
		title_range.Font.Bold = True
		#position picture 5/8" from left, 1.5" from top
		from_left, from_top = 45, 108  #in points
		shape4picture = slide.Shapes.AddPicture(self.file_path, False, True, from_left, from_top)
		#scale to fit 9" x 5"
		w,h = shape4picture.Width, shape4picture.Height
		scalar = min((9*72)/w,(5*72)/h)
		shape4picture.ScaleHeight(scalar,True)
		time.sleep(0.10)
		shape4picture.ScaleWidth(scalar,True)
		#center the picture (assumes standard 10" wide slide)
		shape4picture.Left = (720 - shape4picture.Width)/2
		time.sleep(0.10)


class PptTextPicture(PptTitleSlide):
	'''Slide with picture inserted from file and title.

	:see: http://msdn2.microsoft.com/en-us/library/aa211638(office.11).aspx
	'''
	def __init__(self, text, file_path, title):
		self.file_path = file_path
		self.text = text
		self.title = title
		self.slide = None
		self.layout = ppLayoutTextAndClipart
	def format_content(self):
		slide = self.slide
		title_range = slide.Shapes.Title.TextFrame.TextRange
		title_range.Text = self.title
		title_range.Font.Bold = True
		#
		#add text
		text = self.text
		text_range = slide.Shapes[1].TextFrame.TextRange
		insert_outline(text_range, text)
		#
		#now insert the picture
		shapes = slide.Shapes
		assert len(shapes)==3
		shape = shapes[2]
		assert shape.Type == 14 #Placeholder
		shape4picture = shapes.AddPicture(self.file_path, False, True, shape.Left, shape.Top)
		w,h = shape4picture.Width, shape4picture.Height
		#scale to fit 4" x 5"
		scalar = min((2*72)/w,(2*72)/h)
		shape4picture.ScaleHeight(scalar,True)
		time.sleep(0.10)
		shape4picture.ScaleWidth(scalar,True)
		"""
		print len(shapes)
		print shape4picture, shape4picture.Type,  dir(shape4picture)
		#center the picture (assumes standard 10" wide slide)
		shape4picture.Left = (360 + (360 - shape4picture.Width)/2)
		"""
		time.sleep(0.10)


class PptBasicTable(PptTitleSlide):
	'''Basic PowerPoint table, with optional title.
	If data is R by K, there can be K headers (optional) and R stubs (optional).
	'''
	def __init__(self, data2d, headers=None, stubs=None, title=''):
		self.data = data2d
		self.headers = headers
		self.stubs = stubs
		self.title = title
		self.slide = None
		self.layout = ppLayoutTitleOnly
	def format_content(self):
		slide = self.slide
		#preliminary computations
		data, headers, stubs, title = self.data, self.headers, self.stubs, self.title
		data_nrows, data_ncols = len(data), len(data[0])
		row_offset, col_offset = 1, 1
		table_nrows, table_ncols = data_nrows, data_ncols
		if headers:
			row_offset += 1
			table_nrows +=1
		if stubs:
			col_offset += 1
			table_ncols +=1
		#add shape to hold table
		shape4table = slide.Shapes.AddTable(table_nrows, table_ncols)
		time.sleep(0.1)
		table = shape4table.Table
		#add headers, if any
		if headers:
			for col in range(data_ncols):
				cell = table.Cell(1,col+col_offset)
				cell.Shape.TextFrame.TextRange.Text = str(headers[col])
				time.sleep(0.1)
		#add stubs, if any
		if stubs:
			for row in range(data_nrows):
				cell = table.Cell(row+row_offset,1)
				cell.Shape.TextFrame.TextRange.Text = str(stubs[row])
				time.sleep(0.1)
		#fill in table
		for row in range(data_nrows):
			for col in range(data_ncols):
				cell = table.Cell(row+row_offset,col+col_offset)
				cell.Shape.TextFrame.TextRange.Text = str(data[row][col])
				time.sleep(0.10)
		#format title
		if self.title:
			#add title
			#title_range = slide.Shapes[0].TextFrame.TextRange
			#title_range.Text = self.title
			#title_range.Font.Bold = True
			time.sleep(0.10)
		else: #no title
			slide.Shapes[0].Delete()
			slide.Shapes[0].Top = 20
			slide.Shapes[0].Height = 13*36

