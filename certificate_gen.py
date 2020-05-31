#! python3
# CertGen.py - python module for writing names on certificate templates
#            - can read data from xlsx, csv, tsv, json

# Python stdlib import
from os import getcwd, path

# Pillow module imports
from PIL import Image, ImageDraw, ImageFont

class Certificate():
    """Describes certificate associated properties

    Properties include templateFile, name, fontFile, etc.
    """
    def __init__(
            self, templateFile, name,
            fontFile="src/Merriweather-Bold.ttf",
            fontSize=42, vPosCoordinate=635):
        # should i switch from camel case to lowercase&underscores??
        # templateFile > template_file;    vPosCoordinate > (...) etc.
        self.templateFile = templateFile
        self.name = name
        self.filename = f'Certificate-{"-".join(self.name.split(" "))}.png'
        self.fontFile = fontFile
        self.fontSize = fontSize
        self.vPosCoordinate = vPosCoordinate
        self.fields = [self.templateFile, self.name, self.filename,
                       self.fontFile, self.fontSize, self.vPosCoordinate]

        self.certImg = Image.open(self.templateFile)
        self.drawCert = ImageDraw.Draw(self.certImg)
        self.font = ImageFont.truetype(self.fontFile,
                                       self.fontSize)


    def generate(self, destination=getcwd()):
        '''Generates a certificate image (.png) with name
        and saves the file to destination'''
        self.nameWidth, self.nameHeight = self.drawCert.textsize(
            self.name, self.font)
        self.templateWidth, self.templateHeight = self.certImg.size

        x = (self.templateWidth - self.nameWidth) / 2
        y = self.vPosCoordinate - self.nameHeight

        self.drawCert.text((x, y), self.name, fill='black',
            font=self.font)

        self.certFile = path.join(destination, self.filename)
        self.certImg.save(self.certFile)


    def openImg(self, appPath=r'C:\Windows\System32\mspaint.exe'):
        '''Opens the certificate image for viewing'''
        import subprocess
        self.viewImg = subprocess.Popen([appPath, self.certFile])


    def closeImg(self):
        self.viewImg.terminate()


class Certificates():
    def __init__(self, templateFile, namesFile, file_has_header=False):
        self.templateFile = templateFile
        self.namesFile = namesFile
        self.file_has_header = file_has_header
        self.names = []
        self.filenames = []
        self.filepaths = []
        self.certificates = []


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
            import json
            with open(self.namesFile, 'r') as jsonFile:
                namesDict = json.loads(jsonFile.readlines()[0])
            for value in namesDict.values():
                if isinstance(value, list):
                    for name in value:
                        self.names.append(name)
                else:
                    self.names.append(value)

        for name in self.names:
            certificate = Certificate(templateFile, name)
            self.filenames.append(certificate.filename)



    def generate(self, destination=getcwd()):
        for name in self.names:
            certificate = Certificate(self.templateFile, name)
            certificate.generate(destination)
            self.certificates.append(certificate)
            self.filepaths.append(certificate.certFile)


    def openImgs(self, appPath=r'C:\Windows\System32\mspaint.exe):
        self.openedImgs = []
        for certificate in self.certificates:
            certificate.openImg(appPath)
            self.openedImgs.append(certificate.viewImg)


    def closeImgs(self):
        for viewImg in self.openedImgs:
            viewImg.terminate()


