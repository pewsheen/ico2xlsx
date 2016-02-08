import xlsxwriter
from PIL import Image

# Func Name: rgb2hex
# Func Descripion: Convert RGB to HEX
#
def rgb2hex(r, g, b, a):
	if a == 0:
		r = g = b = 255;
	return '#{:02x}{:02x}{:02x}'.format(r, g, b)

# Open ICO
ICO_Path = raw_input('Enter the path of ICO >> ')

if ICO_Path == '':
	print '[Error] ICO path is empty.'
	exit()

try:
	icon = Image.open(ICO_Path)
except (NameError, AttributeError) as e:
	print '[Error] ICO path is not correct'

print "Image type: " + icon.format
print "Image mode: " + icon.mode
print "Image size: " + str(icon.size)

if icon.format != 'ICO' and icon.format != 'PNG':
	print ICO_Path + ' is not a .ICO file.'
	exit()

pixels = icon.convert('RGBA').load()
width, height = icon.size

# Create xlsx File and Fill Colors
print 'Creating xlsx file...'
workbook = xlsxwriter.Workbook('hello.xlsx')
worksheet = workbook.add_worksheet()

# Set column width to make it square
worksheet.set_column(0, width-1, 1.9)

# Convert ICO to HEX and Fill Cells
print 'Converting...'
for y in range(height):
	for x in range(width):
		r, g, b, a = pixels[x, y]

		format = workbook.add_format()
		format.set_pattern(1)  # This is optional when using a solid fill.
		format.set_bg_color(rgb2hex(r, g, b, a))
		worksheet.write(y, x, '', format)

		print 'x = %s, y = %s, RGBA = %s,%s,%s,%s , hex = %s' % (x, y, r, g, b, a, rgb2hex(r, g, b, a))

# Finish up
workbook.close()
print 'Done!'

