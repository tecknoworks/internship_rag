import os
import zipfile
import xml.etree.ElementTree as ET
import shutil  

class PowerPointXMLExtractor:
    """
    Class to extract text from a .pptx file treated as an XML archive.
    """

    def __init__(self, pptx_file, temp_folder="temp_pptx_extracted"):
        """
        Initializes the extractor with a .pptx file.

        :param pptx_file: Path to the .pptx file.
        :param temp_folder: Temporary folder for extracted files (default: 'temp_pptx_extracted')
        """
        self.pptx_file = pptx_file
        self.temp_folder = temp_folder
        self.namespaces = {
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
            'r': 'http://schemas.openxmlformats.org/package/2006/relationships',
        }
        self.slides = {} 
        self.notes_text = {}
        self.slide_to_notes = {}

    def extract_content(self):
        """
        Extracts content (slides and notes) from the .pptx file.
        """
        if not zipfile.is_zipfile(self.pptx_file):
            raise ValueError("The provided file is not a valid .pptx file.")

        with zipfile.ZipFile(self.pptx_file, 'r') as pptx_zip:
            # Extract slide-to-notes relationships
            self.parse_slide_relationships(pptx_zip)

            # Extract slides
            slide_files = [name for name in pptx_zip.namelist() if name.startswith('ppt/slides/slide') and name.endswith('.xml')]
            for slide_file in slide_files:
                with pptx_zip.open(slide_file) as file:
                    slide_content = file.read()
                    self.parse_slide(slide_content, slide_file)

            # Extract notes
            notes_files = [name for name in pptx_zip.namelist() if name.startswith('ppt/notesSlides/notesSlide') and name.endswith('.xml')]
            for notes_file in notes_files:
                with pptx_zip.open(notes_file) as file:
                    notes_content = file.read()
                    self.parse_notes(notes_content, notes_file)

    def parse_slide_relationships(self, pptx_zip):
        """
        Parses the relationships between slides and notes slides.

        :param pptx_zip: The opened .pptx ZIP archive.
        """
        rels_files = [name for name in pptx_zip.namelist() if name.startswith('ppt/slides/_rels/slide') and name.endswith('.xml.rels')]
        for rels_file in rels_files:
            with pptx_zip.open(rels_file) as file:
                root = ET.parse(file).getroot()
                for rel in root.findall('r:Relationship', self.namespaces):
                    if 'notesSlide' in rel.attrib.get('Type', ''):
                        slide_number = int(os.path.basename(rels_file).replace('slide', '').replace('.xml.rels', ''))
                        notes_number = int(os.path.basename(rel.attrib['Target']).replace('notesSlide', '').replace('.xml', ''))
                        self.slide_to_notes[slide_number] = notes_number

    def parse_slide(self, slide_content, slide_file):
        """
        Parses the XML content of a single slide to extract text.

        :param slide_content: Raw XML content of the slide.
        :param slide_file: Name of the slide file for identification.
        """
        root = ET.fromstring(slide_content)
        title = ""
        description = []

        # XPath to extract text and check size for title
        for p in root.findall('.//a:p', self.namespaces):
            for r in p.findall('a:r', self.namespaces):
                rPr = r.find('a:rPr', self.namespaces)
                t = r.find('a:t', self.namespaces)
                if t is not None and rPr is not None:
                    text_size = rPr.attrib.get('sz')
                    if text_size and int(text_size) >= 2800:  
                        title += t.text.strip() + " "
                    else:
                        description.append(t.text.strip())

        # Extract slide number
        slide_number = int(os.path.basename(slide_file).replace('slide', '').replace('.xml', ''))
        self.slides[slide_number] = {
            "title": title.strip(),
            "description": ' '.join(description),
        }

    def parse_notes(self, notes_content, notes_file):
        """
        Parses the XML content of a single notes slide to extract text.

        :param notes_content: Raw XML content of the notes slide.
        :param notes_file: Name of the notes file for identification.
        """
        root = ET.fromstring(notes_content)
        notes_text = []
        # XPath to extract text from <a:t> elements within <a:p>
        for text_element in root.findall('.//a:p/a:r/a:t', self.namespaces):
            notes_text.append(text_element.text.strip())

        # Store the notes text
        notes_number = int(os.path.basename(notes_file).replace('notesSlide', '').replace('.xml', ''))
        self.notes_text[notes_number] = ' '.join(notes_text)

    def display_content(self):
        """
        Displays the extracted slides in the desired format.
        """
        print("Slides Content:")
        for slide_number in sorted(self.slides.keys()):
            slide = self.slides[slide_number]
            print(f"Slide number: {slide_number}")
            print(f"  Title: {slide['title']}")
            print(f"  Description: {slide['description']}")

            # Display the corresponding notes if available
            notes_number = self.slide_to_notes.get(slide_number)
            if notes_number and notes_number in self.notes_text:
                print(f"  Notes: {self.notes_text[notes_number]}")
            print()

    def cleanup(self):
        """
        Deletes the temporary folder and its contents after processing.
        """
        if os.path.exists(self.temp_folder):
            shutil.rmtree(self.temp_folder)
            print(f"Temporary folder '{self.temp_folder}' has been removed.")


if __name__ == "__main__":
    pptx_path = os.path.join("data", "path to pptx")

    try:
        extractor = PowerPointXMLExtractor(pptx_path)
        extractor.extract_content()
        extractor.display_content()

    except Exception as e:
        print(f"Error: {e}")
    
    finally:
        extractor.cleanup()
