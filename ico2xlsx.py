import xlsxwriter
import operator
import struct
from PIL import BmpImagePlugin, PngImagePlugin, Image

def rgb2hex(r, g, b, a):
	# Some transparent pixel reads RGBA:0,0,0,0
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

'''
	Some ICO file is not really ICO follow Windows Spec, IT'S PNG!!!
'''
def loadPNG(imgPath):
	try:
		img = Image.open(imgPath)
	except (NameError, AttributeError) as e:
		print '[Error] ICO path is not correct'
		exit()

	if img.format != 'ICO' and img.format != 'PNG':
		print imgPath + ' is not a valid filetype.'
		exit()

	return img, img.size

def load_icon(imgPath, index=None):
	if isinstance(imgPath, basestring):
		file = open(imgPath, 'rb')

	try:
		header = struct.unpack('<3H', file.read(6))
	except:
		raise IOError('Not an ICO file')

	# Check magic
	try:
		if header[:2] != (0, 1):
			raise IOError('Not an ICO file')
	except:
		return loadPNG(imgPath)

	# Collect icon directories
	directories = []
	for i in xrange(header[2]):
		directory = list(struct.unpack('<4B2H2I', file.read(16)))
		for j in xrange(3):
			if not directory[j]:
				directory[j] = 256

		directories.append(directory)

	if index is None:
		# Select best icon
		directory = max(directories, key=operator.itemgetter(slice(0, 3)))
	else:
		directory = directories[index]

	# Seek to the bitmap data
	file.seek(directory[7])

	prefix = file.read(16)
	file.seek(-16, 1)

	if PngImagePlugin._accept(prefix):
		# Windows Vista icon with PNG inside
		image = PngImagePlugin.PngImageFile(file)
	else:
		# Load XOR bitmap
		image = BmpImagePlugin.DibImageFile(file)
		if image.mode == 'RGBA':
			# Windows XP 32-bit color depth icon without AND bitmap
			pass
		else:
			# Patch up the bitmap height
			image.size = image.size[0], image.size[1] >> 1
			d, e, o, a = image.tile[0]
			image.tile[0] = d, (0, 0) + image.size, o, a

			# Calculate AND bitmap dimensions. See
			offset = o + a[1] * image.size[1]
			stride = ((image.size[0] + 31) >> 5) << 2
			size = stride * image.size[1]

			# Load AND bitmap
			file.seek(offset)
			string = file.read(size)
			mask = Image.fromstring('1', image.size, string, 'raw',
									('1;I', stride, -1))

			image = image.convert('RGBA')
			image.putalpha(mask)

	return image, image.size

if __name__ == '__main__':
	# Open ICO
	ICO_Path = raw_input('Enter the path of ICO >> ')

	if ICO_Path == '':
		print '[Error] ICO path is empty.'
		exit()

	img, (width, height) = load_icon(ICO_Path)
	print "Image type: " + str(img.format)
	print "Image mode: " + str(img.mode)
	print "Image size: " + str(img.size)
	print "Image band: " + str(img.getbands())
	pixels = img.convert('RGBA').load()

	# Create xlsx File
	print 'Creating xlsx file...'
	workbook, worksheet = createXlsx('favicon.xlsx')

	# Set column width to make it square
	worksheet.set_column(0, width-1, 2.4)

	# Convert ICO to HEX and Fill Cells
	print 'Converting...'
	convert(width, height, pixels, workbook, worksheet)

	# Finish up
	workbook.close()

	print 'Done!'
