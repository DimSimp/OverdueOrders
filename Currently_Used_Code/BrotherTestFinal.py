

try:

    from pil import Image, ImageDraw, ImageFont
    
except ModuleNotFoundError as err:
    try:
        from PIL import Image, ImageDraw, ImageFont

    except:
        from pil import Image, ImageDraw, ImageFont
import pyperclip
import brother_ql
from brother_ql.raster import BrotherQLRaster
from brother_ql.backends.helpers import send


filename = r'C:\Python\test2.png' #this will vary system to system

img = Image.new('RGB', (400, 240), color=(255, 255, 255))

fnt = ImageFont.truetype(r'C:\Windows\Fonts\calibri.ttf', 23, encoding='unic')

text = pyperclip.paste()
text = text.replace("Purchase Order #	", "")
text = text.replace("Date Required	-", "")
text = text.replace("Name	", "Name: ")
text = text.replace("Company	", "")
text = text.replace("Address	", "Address: \n")
text = text.replace("City	", "")
text = text.replace("State	", "")
text = text.replace("Postal Code	", "")
text = text.replace("Country	", "")
text = text.replace("Phone	", "Phone: ")
text = text.replace("	", "")
text = text.replace("Australia (AU)", "")
text = text.replace("Fax", "")
text = text.replace("Shipping Address\r\n \r\n\r\n", "")
text = text.replace("\r\n\r\nName:", "Name:")




d = ImageDraw.Draw(img)
d.multiline_text((10,10), text, font=fnt, fill=(0,0,0))
img.save(filename)

colorImage = Image.open(filename)
rotatedimage = colorImage.transpose(Image.ROTATE_90)
rotatedimage.save(filename)

######### PART 2 NOW WE HAVE IMAGE TO PRINT

PRINTER_IDENTIFIER = brother_ql.backends.helpers.discover('pyusb')

PRINTER_IDENTIFIER = brother_ql.backends.helpers.discover('pyusb')[0]['identifier']

if brother_ql.backends.helpers.discover('pyusb')[0]['identifier'] == 'usb://0x04f9:0x2042':
    PRINTER_IDENTIFIER = 'usb://0x04F9:0x2042'

    printer = BrotherQLRaster('QL-700')
    printermodel = 'QL-700'

# if brother_ql.backends.helpers.discover('pyusb')[0]['identifier'] == 'usb://0x04f9:0x2042':
#     PRINTER_IDENTIFIER = 'usb://0x04F9:0x2042'
#
#     printer = BrotherQLRaster('QL-700')
#     printermodel = 'QL-700'

elif brother_ql.backends.helpers.discover('pyusb')[0]['identifier'] == 'usb://0x04f9:0x2042_Љ':
    PRINTER_IDENTIFIER = 'usb://0x04F9:0x2042'

    printer = BrotherQLRaster('QL-700')
    printermodel = 'QL-700'

elif brother_ql.backends.helpers.discover('pyusb')[0]['identifier'] == 'usb://0x04f9:0x2028':
    PRINTER_IDENTIFIER = 'usb://0x04F9:0x2028'

    printer = BrotherQLRaster('QL-570')
    printermodel = 'QL-570'

elif brother_ql.backends.helpers.discover('pyusb')[0]['identifier'] == 'usb://0x04f9:0x2028_Љ':
    PRINTER_IDENTIFIER = 'usb://0x04F9:0x2028'

    printer = BrotherQLRaster('QL-570')
    printermodel = 'QL-570'
def sendToPrinter(PRINTER_IDENTIFIER, printer):
    print_data = brother_ql.brother_ql_create.convert(printer, [filename], '62', dither=True)
    send(print_data, PRINTER_IDENTIFIER)


try:

    sendToPrinter(PRINTER_IDENTIFIER, printer)

except:

    PRINTER_IDENTIFIER = 'usb://0x04F9:0x2028'

    printer = BrotherQLRaster('QL-570')

    sendToPrinter(PRINTER_IDENTIFIER, printer)