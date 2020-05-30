#! python3
# reads name from xlsx, csv, tsv, json files

import os
from PIL import Image, ImageDraw, ImageFont

class Certificate():
    def __init__(
            self, templateFile, name,
            fontFile="src/Merriweather-Bold.ttf",
            fontSize=42, vPosCoordinate=635):
        self.templateFile = templateFile
        self.name = name
        self.filename = f'Certificate-{"-".join(self.name.split(" "))}.png'
        self.fontFile = fontFile
        self.fontSize = fontSize
        self.vPosCoordinate = vPosCoordinate

        self.certImg = Image.open(self.templateFile)
        self.drawCert = ImageDraw.Draw(self.certImg)
        self.font = ImageFont.truetype(self.fontFile,
                                       self.fontSize)


    def generate(self, destination=os.getcwd()):
        self.nameWidth, self.nameHeight = self.drawCert.textsize(
            self.name, self.font)
        self.templateWidth, self.templateHeight = self.certImg.size

        self.x = (self.templateWidth - self.nameWidth) / 2
        self.y = self.vPosCoordinate - self.nameHeight

        self.drawCert.text((self.x, self.y), self.name, fill='black',
            font=self.font)

        self.certFile = os.path.join(destination, self.filename)
        self.certImg.save(self.certFile)


    def openImg(self, appPath=r'C:\Windows\System32\mspaint.exe'):
        from subprocess import Popen
        return Popen([appPath, self.certFile])


class Certificates():
    def __init__(self, namesFile, file_has_header=False):
        self.namesFile = namesFile
        self.file_has_header = file_has_header
        self.certificates = []
        self.names = []

        if self.namesFile.endswith('.xlsx'):
            #TODO: read different formats of tables
            import openpyxl
            excelFile = openpyxl.open(self.namesFile)
            sheet = excelFile.active
            for cell in sheet.iter_rows():
                self.names.append(cell[0].value)

        elif self.namesFile.endswith('.csv') or self.namesFile.endswith('.tsv'):
            import csv
            with open(self.namesFile, 'r') as csvFile:
                if self.file_has_header:
                    csvReader = csv.DictReader(csvFile)
                    fieldnames = csvReader.fieldnames
                    for row in csvReader:
                        for name in row.values():
                            self.names.append(name)
                else:
                    csvReader = csv.reader(csvFile)
                    for row in csvReader:
                        if self.namesFile.endswith('.tsv'):
                            if len(row) > 0:
                                for name in row[0].split('\t'):
                                    self.names.append(name)
                        else:
                            for name in row:
                                self.names.append(name)
        elif self.namesFile.endswith('.json'):
            pass

# for file in os.listdir('src'):
#     if file.endswith('csv') or file.endswith('tsv') or file.endswith('xlsx'):
#         certificates = Certificates(os.path.join('src', file))
#         certificate = Certificate('src/certificate_template_1.png', certificates.names[0])
#         certificate.generate()
#         certificate.openImg()
#         print(certificate.templateFile)
#         break


