import os
from PyPDF2 import PdfReader

class PDFTextExtractor:
    """
    A class for extracting text from PDF files using PyPDF2.
    """

    def __init__(self, file_path):
        """
        Initialize the extractor with the path to the PDF file.
        
        :param file_path: Path to the PDF file.
        """
        self.file_path = file_path
        self.validate_file()

    def validate_file(self):
        """
        Validate if the file exists and has a .pdf extension.
        """
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"The file {self.file_path} does not exist.")
        if not self.file_path.lower().endswith('.pdf'):
            raise ValueError("The provided file is not a PDF.")

    def extract_text(self):
        """
        Extract text from the PDF file.

        :return: A string containing the extracted text.
        """
        try:
            reader = PdfReader(self.file_path)
            text = ""
            for page in reader.pages:
                text += page.extract_text() + "\n"
            return text.strip()
        except Exception as e:
            raise RuntimeError(f"Failed to extract text from PDF: {e}")

def main():
    """
    Main function to demonstrate PDF text extraction.
    """
    # Modify this to the path of your PDF file
    pdf_path = os.path.join("data", "test_doc.pdf")

    try:
        extractor = PDFTextExtractor(pdf_path)
        extracted_text = extractor.extract_text()

        # Save extracted text to a file (optional)
        output_path = os.path.splitext(pdf_path)[0] + "_extracted.txt"
        with open(output_path, 'w', encoding='utf-8') as output_file:
            output_file.write(extracted_text)
        
        print(f"Text successfully extracted and saved to {output_path}")

    except (FileNotFoundError, ValueError, RuntimeError) as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()