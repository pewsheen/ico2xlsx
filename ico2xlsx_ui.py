import xlsxwriter
from PIL import Image
from Tkinter import Tk
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

def loadICO(imgPath):
	try:
		img = Image.open(imgPath)
	except (NameError, AttributeError) as e:
		print '[Error] ICO path is not correct'
		exit()

	print "Image type: " + img.format
	print "Image mode: " + img.mode
	print "Image size: " + str(img.size)

	if img.format != 'ICO' and img.format != 'PNG':
		print imgPath + ' is not a valid path.'
		exit()

	return img, img.size

def convert(width, height, pixels, workbook, worksheet):
	for y in range(height):
		for x in range(width):
			r, g, b, a = pixels[x, y]

			format = workbook.add_format()
			format.set_pattern(1)  # This is optional when using a solid fill.
			format.set_bg_color(rgb2hex(r, g, b, a))
			worksheet.write(y, x, '', format)
			print 'x = %s, y = %s, RGBA = %s,%s,%s,%s , hex = %s' % (x, y, r, g, b, a, rgb2hex(r, g, b, a))

# Open ICO
Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
ICO_Path = askopenfilename() # show an "Open" dialog box and return the path to the selected file

if ICO_Path == '':
	print '[Error] ICO path is empty.'
	exit()

icon, (width, height) = loadICO(ICO_Path)
pixels = icon.convert('RGBA').load()

# Create xlsx File
print 'Creating xlsx file...'
workbook, worksheet = createXlsx('hello.xlsx')

# Set column width to make it square
worksheet.set_column(0, width-1, 1.9)

# Convert ICO to HEX and Fill Cells
print 'Converting...'
convert(width, height, pixels, workbook, worksheet)

# Finish up
workbook.close()

print 'Done!'
