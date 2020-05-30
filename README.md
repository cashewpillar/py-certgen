# Certificate-Generator
Desktop application for generating certificates from a list of names (file formats: .xlsx .csv .tsv .json)

###Certificate class
Create a certificate object with fields that provide information such as:
    1. the templateFile (certificate img) to be used
    2. name to be written on template
    3. filename to be saved
    4. fontfile to use & fontsize
    5. vertical coordinate to place the name on the template
With methods:
    1. generate(destination=defaultsToCurrentDirectory) - generates certificate and saves (destination defaults to current dir)
    2. openImg(appPath=pathnameForAppToUse) - opens generated certificate (appPath defaults to mspaint.exe)
  
