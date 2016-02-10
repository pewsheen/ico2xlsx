import xlsxwriter
import PIL.Image
from Tkinter import *
from tkFileDialog import askopenfilename

# Func Name: rgb2hex
# Func Descripion: Convert RGB to HEX
#
def rgb2hex(r, g, b, a):
	if a == 0:
		r = g = b = 255;
	return '#{:02x}{:02x}{:02x}'.format(r, g, b)

def createXlsx(xlsxPath):
	workbook = xlsxwriter.Workbook(xlsxPath)
	worksheet = workbook.add_worksheet()

	return workbook, worksheet

def convert(width, height, pixels, workbook, worksheet):
	for y in range(height):
		for x in range(width):
			r, g, b, a = pixels[x, y]

			format = workbook.add_format()
			format.set_pattern(1)  # This is optional when using a solid fill.
			format.set_bg_color(rgb2hex(r, g, b, a))
			worksheet.write(y, x, '', format)
			print 'x = %s, y = %s, RGBA = %s,%s,%s,%s , hex = %s' % (x, y, r, g, b, a, rgb2hex(r, g, b, a))

def loadICO(_imgPath):
	try:
		img = PIL.Image.open(_imgPath)
	except (NameError, AttributeError) as e:
		print '[Error] ICO path is not correct'
		# exit()
	except IOError as e:
		if str(e).find('cannot identify image file') != -1:
			print _imgPath + ' is not a valid path.'
			# exit()

	print "Image type: " + img.format
	print "Image mode: " + img.mode
	print "Image size: " + str(img.size)

	if img.format != 'ICO' and img.format != 'PNG':
		print _imgPath + ' is not a valid path.'
		# exit()

	return img, img.size

class GUIDemo(Frame):
	imgPath = ''

	def __init__(self, master=None):
		Frame.__init__(self, master)
		self.grid()
		self.createWidgets()
 
	def createWidgets(self):
		self.inputText = Label(self)
		self.inputText["text"] = "ICO Path:"
		self.inputText.grid(row=0, column=0)
		self.inputField = Entry(self)
		self.inputField["width"] = 50
		self.inputField.grid(row=0, column=1, columnspan=6)
 
		# self.outputText = Label(self)
		# self.outputText["text"] = "Output:"
		# self.outputText.grid(row=1, column=0)
		# self.outputField = Entry(self)
		# self.outputField["width"] = 50
		# self.outputField.grid(row=1, column=1, columnspan=6)
		 
		self.load = Button(self)
		self.load["text"] = "Load"
		self.load.grid(row=2, column=5)
		self.load["command"] =  self.loadMethod
		
		self.convert = Button(self)
		self.convert["text"] = "Convert"
		self.convert.grid(row=2, column=6)
		self.convert["command"] =  self.convertMethod
 
		self.displayText = Label(self)
		self.displayText["text"] = ""
		self.displayText.grid(row=3, column=0, columnspan=7)
	  
	def loadMethod(self):
		self.imgPath = askopenfilename()
		self.inputField.insert(END, self.imgPath) 
 
	def convertMethod(self):
		self.imgPath = self.inputField.get()
		print 'imgPath = ' + self.imgPath
		print 'entry = ' + self.inputField.get()
		self.displayText["text"] = self.imgPath

		if self.imgPath == '':
			self.displayText["text"] = '[Error] ICO path is empty.'

		icon, (width, height) = loadICO(self.imgPath)
		pixels = icon.convert('RGBA').load()

		# Create xlsx File
		print 'Creating xlsx file...'
		self.displayText["text"] = 'Creating xlsx file...'
		workbook, worksheet = createXlsx('hello.xlsx')

		# Set column width to make it square
		worksheet.set_column(0, width-1, 1.9)

		# Convert ICO to HEX and Fill Cells
		print 'Converting...'
		self.displayText["text"] = 'Converting...'
		convert(width, height, pixels, workbook, worksheet)

		# Finish up
		workbook.close()

		print 'Done!'
		self.displayText["text"] = 'Done!'
		# exit()

if __name__ == '__main__':
	
	# Open ICO
	root = Tk()
	root.title("ico2xlsx Converter")
	app = GUIDemo(master=root)
	app.mainloop()
