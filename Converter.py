from pythonnet import load
load("coreclr")
import clr
import os


# ASPOSE_DLL_DIRECTORY = r"C:\Users\hanih\Documents\yellowsys\aspose\publish"  # Replace with your directory path
# ASPOSE_DLL_DIRECTORY=r"C:\Users\hanih\Documents\yellowsys\aspose\publish"
ASPOSE_DLL_DIRECTORY=os.path.join(os.getcwd(),"net7.0","win-x64","publish")
dlls = [
    "Aspose.Words.dll",
    "Aspose.Slides.dll",
]

for dll in dlls:
    clr.AddReference(os.path.join(ASPOSE_DLL_DIRECTORY, dll))


from Aspose.Words import Document as WordsDocument, License as WordsLicense, SaveFormat as WordsSaveFormat
from Aspose.Slides import Presentation, License as SlidesLicense, Export

class Converter:
    def __init__(self, license_path):
        self._license_path = license_path
        self.apply_license()

    def apply_license(self):
        # Words
        WordsLicense().SetLicense(self._license_path)

        # Slides
        SlidesLicense().SetLicense(self._license_path)

       
    def convert_to_pdf(self, input_path, output_path):
        if not input_path or not output_path:
            raise ValueError("Input or output path is null or empty.")

        if not os.path.exists(input_path):
            raise FileNotFoundError(f"File not found: {input_path}")

        file_extension = os.path.splitext(input_path)[1].lower()

        if file_extension in [".doc", ".docx"]:
            doc = WordsDocument(input_path)
            doc.Save(output_path, WordsSaveFormat.Pdf)

        elif file_extension in [".ppt", ".pptx"]:
            pres= Presentation(input_path)
            pres.Save(output_path, Export.SaveFormat.Pdf)

        else:
            raise ValueError(f"File format {file_extension} is not supported.")
        
