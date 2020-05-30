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
  
