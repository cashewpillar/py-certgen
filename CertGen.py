#! python3
# reads name from xlsx, csv, json files

import os
import openpyxl
from PIL import Image, ImageDraw, ImageFont

# SOURCE = 'C:/python_files/certificate_generator/src'
# DESTINATION = 'C:/python_files/certificate_generator/generated_certificates'
# CERTIFICATE_TEMPLATE = 'certificate_template_1.png'
# names_xlsx = 'certificate_names.xlsx'

# os.chdir(SOURCE)
# FONT = ImageFont.truetype('Merriweather-Bold.ttf', 42)


# generate_certificates(names_xlsx)



# May 29 Trial
# Aligning name to certificate template

# Set templateWidth and templateHeight
# Set a vertical placement for name
# Get nameWidth and nameHeight
# Subtract nameWidthHalf from templateWidthHalf
# Subtract nameHeight from (templateHeight - verticalPlacementHeight)

import subprocess
# xl = openpyxl.open('certificate_names.xlsx')
# sheet = xl.active

# for cell in sheet.iter_rows():
#     filename = generateCert(cell[0].value)
#     openImg(os.path.join(r'C:\python_files\certificate_generator\0529_test', filename))

os.chdir('src')

# FONT = ImageFont.truetype('Merriweather-Bold.ttf', 42)
class Certificate():
    def __init__(
            self, template, name,
            fontFamily="Merriweather-Bold.ttf",
            fontSize=42, vPosCoordinate=635):
        self.template = template
        self.name = name
        self.filename = f'Certificate-{"-".join(self.name.split(" "))}.png'
        self.fontFamily = fontFamily
        self.fontSize = fontSize
        self.vPosCoordinate = vPosCoordinate

        self.certImg = Image.open(self.template)
        self.drawCert = ImageDraw.Draw(self.certImg)
        self.font = ImageFont.truetype(self.fontFamily,
                                       self.fontSize)


    def generate(self, destination=os.getcwd()):
        self.nameWidth, self.nameHeight = self.drawCert.textsize(
            self.name, self.font)
        self.templateWidth, self.templateHeight = self.certImg.size

        self.x = (self.templateWidth - self.nameWidth) / 2
        self.y = self.vPosCoordinate - self.nameHeight

        self.drawCert.text((self.x,self.y), self.name, fill='black',
            font=self.font)

        self.certFile = os.path.join(destination, self.filename)
        self.certImg.save(self.certFile)


    def openImg(self, appPath=r'C:\Windows\System32\mspaint.exe'):
        return subprocess.Popen([appPath, self.certFile])


