from pythonnet import load
load("coreclr")

import clr
import os


publish_dir=r"C:\Users\hanih\Documents\yellowsys\convertor-final\convertor-final\bin\Release\net7.0\win-x64"
clr.AddReference(os.path.join(publish_dir, "convertor-final.dll"))

dir=os.path.join(publish_dir, "publish")
clr.AddReference(os.path.join(dir, "Aspose.Words.dll"))
clr.AddReference(os.path.join(dir, "Aspose.Slides.dll"))

from DocumentConversion import Converter

license_path = r"C:\Users\hanih\Documents\yellowsys\aspose\Aspose.TotalProductFamily.lic"
converter = Converter(license_path)


input_path_pptx = r"C:\Users\hanih\Documents\yellowsys\aspose\9anoungypt.pptx"
# output_path_pdf = r"C:\Users\hanih\Documents\yellowsys\aspose\9anoungypt.pdf"
output_path_pdf = r"9anoungypt.pdf"

converter.ConvertToPdf(input_path_pptx, output_path_pdf)
converter.ConvertToPdf("mybestdocument.docx", "mybestdocument.pdf" )
