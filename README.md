# RAG Project

## Overview

This project consists of two Python scripts, `01_pptx.py` and `02_pdf.py`, that are designed to extract text from PowerPoint (`.pptx`) and PDF (`.pdf`) files, respectively. The extracted text can be displayed or saved into a text file for further processing. Both scripts read files from the `data` directory, which contains sample files (`test_doc.pdf` and `test_pres.pptx`) for testing.

## Features

- **`01_pptx.py`**: Extracts text from PowerPoint presentations, including slides and speaker notes.
- **`02_pdf.py`**: Extracts text from PDF files.

## Prerequisites

Before running the scripts, make sure you have the following installed:

- **Python 3.x** (preferably Python 3.7 or higher)
- **pip**: The Python package installer

The required dependencies for both scripts are listed in the `requirements.txt` file. 

## Setup and Installation

1. **Clone the Repository**: 

   If you haven't already, clone this repository to your local machine:

   ```bash
   git clone https://github.com/tecknoworks/internship_rag
   cd internship_rag
   ```

2. **Install Dependencies**: 

   You can install all required dependencies by running the following command:

   ```bash
   pip install -r requirements.txt
   ```

3. **Place the Test Files**:

   Ensure that the test files `test_doc.pdf` and `test_pres.pptx` are located in the `data` folder. The folder structure should look like this:

   ```
   /internship_rag
   ├── 01_pptx.py
   ├── 02_pdf.py
   ├── data/
   │   ├── test_doc.pdf
   │   └── test_pres.pptx
   ├── requirements.txt
   └── README.md
   ```

## Usage

### Running the PowerPoint Extraction Script (`01_pptx.py`)

This script extracts text from a PowerPoint file (`.pptx`). It extracts the content of the slides and associated speaker notes. You can customize the PowerPoint file path in the script.

1. Open the `01_pptx.py` script.
2. Ensure the `pptx_path` points to the correct `.pptx` file within the `data` folder:

   ```python
   pptx_path = os.path.join("data", "test_pres.pptx")
   ```

3. Run the script:

   ```bash
   python 01_pptx.py
   ```

   The script will extract the slide text and speaker notes, display them in the terminal, and return the extracted data in the form of dictionaries. It also performs cleanup by removing the temporary folder used for extraction.

### Running the PDF Extraction Script (`02_pdf.py`)

This script extracts text from a PDF file (`.pdf`). It reads the file, extracts all text from each page, and optionally saves it to a text file.

1. Open the `02_pdf.py` script.
2. Ensure the `pdf_path` points to the correct `.pdf` file in the `data` folder:

   ```python
   pdf_path = os.path.join("data", "test_doc.pdf")
   ```

3. Run the script:

   ```bash
   python 02_pdf.py
   ```

   The script will extract text from the PDF and save it as a `.txt` file in the same directory as the PDF. The extracted text will also be printed to the terminal.

## File Descriptions

- **`01_pptx.py`**: A script for extracting text from PowerPoint presentations, including both slides and speaker notes. The extracted text is printed to the console.
- **`02_pdf.py`**: A script for extracting text from PDF files. The extracted text is saved to a `.txt` file and printed to the console.
- **`requirements.txt`**: Contains all the necessary Python libraries to run the scripts. It includes dependencies like `PyPDF2` for PDF text extraction.
- **`data/`**: Folder containing the test files, `test_doc.pdf` and `test_pres.pptx`.

## Requirements

The following Python libraries are required to run the scripts:

- **`PyPDF2`**: A library for reading PDF files and extracting text from them.
- **`shutil`**: A built-in Python library for file operations (used in `01_pptx.py` for cleanup).
- **`zipfile`** and **`xml.etree.ElementTree`**: Built-in Python libraries for extracting content from PowerPoint `.pptx` files.

To install the required libraries, you can run:

```bash
pip install -r requirements.txt
```

## Error Handling

Both scripts include error handling for common issues, such as:

- **FileNotFoundError**: If the specified file does not exist.
- **ValueError**: If the file is not of the expected type (i.e., not a `.pdf` or `.pptx`).
- **RuntimeError**: In case of issues during text extraction.

Error messages will be printed to the console, and the script will attempt to handle the issue gracefully.

## Cleanup

Both scripts perform cleanup tasks:

- **`01_pptx.py`**: Deletes the temporary folder used for extracting PowerPoint content (`temp_pptx_extracted`) after the script finishes running.
- **`02_pdf.py`**: Does not require cleanup, but any extracted text will be saved to a `.txt` file.

## Contributing

If you'd like to contribute to this project, feel free to fork the repository and create a pull request with your improvements. Make sure to:

- Follow the existing code style.
- Write clear commit messages.
- Provide test cases for new features or bug fixes.

## Challenges and Solutions

### 1. **Extracting Titles**

One of the main challenges encountered during the development of the PowerPoint extraction functionality (`01_pptx.py`) was correctly associating slide notes with their corresponding slides. PowerPoint presentations store notes in separate XML files located in the `ppt/notesSlides` folder, while slide content is stored in the `ppt/slides` folder. 

Initially, it was difficult to properly link notes to the correct slides, especially when the filenames for notes (e.g., `notesSlide1.xml`) didn’t directly match the slide numbers (e.g., `slide3.xml`). The core of the issue was the mismatch between the naming conventions of slides and notes in the PowerPoint file structure.

### Solution:

To overcome this, we utilized the relationship mapping provided within the PowerPoint `.pptx` structure. The `slide_to_notes` dictionary was used to map the slide files (e.g., `slide3.xml`) to their corresponding notes files (e.g., `notesSlide1.xml`).

### 2. **Linking Notes to Slides** 

Another challenge was accurately identifying and extracting slide titles from the XML content. Titles in PowerPoint slides are not explicitly labeled as such in the XML files, and they often appear alongside other text elements, making it difficult to distinguish them from regular content.

### Solution:
We identified titles based on their font size. In the XML structure of PowerPoint slides, titles are typically formatted with a larger font size compared to other text elements. By checking the sz attribute of text elements and extracting only those with a font size of 28 or greater, we ensured that the extracted text was most likely the title of the slide. 

### 3. **Linking Notes to Slides**
The order of slides in the extracted content was initially inconsistent due to the way files are stored in the .pptx archive. Slide files (e.g., slide1.xml, slide2.xml) are not always processed in numerical order, leading to unordered results.

### Solution:
To maintain the correct order, we extracted the slide number from each file name (e.g., extracting 1 from slide1.xml) and stored the slides in a dictionary with the slide number as the key. Before displaying the content, we sorted the dictionary by slide number, ensuring that the output matched the natural order of slides in the presentation.

## Troubleshooting

If you encounter any issues, here are a few things to check:

1. **Missing files**: Ensure that the test files (`test_doc.pdf` and `test_pres.pptx`) are present in the `data/` folder.
2. **Dependencies**: Make sure all required libraries are installed by running `pip install -r requirements.txt`.
3. **File type errors**: Double-check that the file paths point to valid `.pptx` and `.pdf` files in the `data/` folder.

If you're still having trouble, feel free to open an issue on the GitHub repository.
