try:

    from pil import Image, ImageDraw, ImageFont

except ModuleNotFoundError as err:
    from PIL import Image, ImageDraw, ImageFont
import pyperclip
import brother_ql
from brother_ql.raster import BrotherQLRaster
from brother_ql.backends.helpers import send
import re, openpyxl, csv

csv_file = rf"\\SERVER\Python\MasterPickSheet.csv"

reader = csv.reader(open(csv_file, encoding='utf-8'), delimiter=',')

for row in reader:
    if row[3] == "2) Back Area":

        if 'pick' in row[0].lower() or 'plectrum' in row[0].lower() or 'guitar bridge pin' in row[0].lower():
            pass
        elif 'player' in row[0].lower():
            continue
        else:
            continue

        for _ in range(int(row[2])):

            #print(row)

            text = f"[SKU: {row[1]}] {row[0]}"


            text = re.sub("(.{30})", "\\1\n", text, 0, re.DOTALL)

            nlines = text.count('\n')

            image_width = 22*(nlines+1)


            filename = r'C:\Python\test2.png'  # this will vary system to system

            img = Image.new('RGB', (240, image_width), color=(255, 255, 255))

            fnt = ImageFont.truetype(r'C:\Windows\Fonts\calibri.ttf', 18, encoding='unic')

            d = ImageDraw.Draw(img)
            d.multiline_text((10 ,10), text, font=fnt, fill=(0 ,0 ,0))
            img.save(filename)

            # colorImage = Image.open(filename)
            # rotatedimage = colorImage.transpose(Image.ROTATE_90)
            # rotatedimage.save(filename)

            ######### PART 2 NOW WE HAVE IMAGE TO PRINT

            PRINTER_IDENTIFIER = brother_ql.backends.helpers.discover('pyusb')[0]['identifier']

            if brother_ql.backends.helpers.discover('pyusb')[0]['identifier'] == 'usb://0x04f9:0x2042':
                PRINTER_IDENTIFIER = 'usb://0x04F9:0x2042'

                printer = BrotherQLRaster('QL-700')
                printermodel = 'QL-700'

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