# Certificate-Generator
Desktop application for generating certificates from a list of names (file formats: .xlsx .csv .tsv .json)

### Certificate class
##### Create a certificate object with fields that provide information such as:
* the templateFile (certificate img) to be used
* name to be written on template
* filename to be saved
* fontfile to use & fontsize
* vertical coordinate to place the name on the template
##### With methods:
* generate(destination=defaultsToCurrentDirectory) - generates certificate and saves (destination defaults to current dir)
* openImg(appPath=pathnameForAppToUse) - opens generated certificate (appPath defaults to mspaint.exe)
##### Sample:
```python
import CertGen
certObj = Certificate('src/certificate_template_1.png', 'Draco Malfoy') # create a Certificate object
print(certObj.templateFile) # access certificate template filename
print(certObj.name) # name
print(certObj.filename) # processed name for saving file
print(certObj.fontFile)
print(certObj.fontSize)
print(certObj.vPosCoordinate) # vertical alignment for positioning name into template
print(certObj.fields) # view fields on certificate object
certObj.generate() # generate certificate file and save
certObj.openImg() # open generated certificate file
```


### Certificates class
#TODO - multiple certificate generation
