import xlsxwriter
from PIL import Image

# Func Name: rgb2hex
# Func Descripion: Convert RGB to HEX
#
def rgb2hex(r, g, b, a):
	if a == 0:
		r = g = b = 255;
	return '#{:02x}{:02x}{:02x}'.format(r, g, b)

# Create xlsx File and Fill Colors
workbook = xlsxwriter.Workbook('hello.xlsx')
worksheet = workbook.add_worksheet()

# Load ICO File and Get Size
icon = Image.open("favicon.ico")

# print "Image type: " + icon.format
# print "Image mode: " + icon.mode
# print "Image size: " + str(icon.size)

pixels = icon.convert('RGBA').load()
width, height = icon.size

worksheet.set_column(0, width-1, 1.9)

# Convert ICO to HEX and Fill Cells
for y in range(height):
	for x in range(width):
		r, g, b, a = pixels[x, y]

		format = workbook.add_format()
		format.set_pattern(1)  # This is optional when using a solid fill.
		format.set_bg_color(rgb2hex(r, g, b, a))
		worksheet.write(y, x, '', format)

		# print 'x = %s, y = %s, RGBA = %s,%s,%s,%s , hex = %s' % (x, y, r, g, b, a, rgb2hex(r, g, b, a))

workbook.close()

