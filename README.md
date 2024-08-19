# ConvertPDFtoDOCX

A Python script that converts a PDF to a Word document while performing OCR with font and layout recognition.

## Features

- Converts PDF files to Word documents (.docx)
- Performs Optical Character Recognition (OCR) to extract text from images
- Maintains font and layout recognition
- Handles multi-page PDFs and adds appropriate page breaks in the Word document
- Customizable OCR configurations

## Requirements

- Python 3.x
- Required Python packages:
  - `pdf2image`
  - `pytesseract`
  - `python-docx`
  - `Pillow`
  - `opencv-python`
  - `numpy`
- Poppler (for PDF to image conversion)

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/ConvertPDFtoDOCX.git
    cd ConvertPDFtoDOCX
    ```

2. Install the required Python packages:
    ```sh
    pip install -r requirements.txt
    ```

3. Install Poppler:
    - **Windows**: 
        1. Download Poppler for Windows: [Poppler for Windows](http://blog.alivate.com.au/poppler-windows/)
        2. Extract the downloaded file and add the `bin` folder to your system's PATH.
    - **macOS**:
        1. Install Homebrew if not already installed: [Homebrew](https://brew.sh/)
        2. Run: `brew install poppler`
    - **Linux**:
        1. Run: `sudo apt-get install poppler-utils`

## Usage

1. Place the PDF file you want to convert in the `pdf` directory.
2. Run the script:
    ```sh
    python pdf-to-word-ocr-converter.py
    ```

3. The converted Word document will be saved in the `output` directory.

## Example

To convert a PDF named `input.pdf` to a Word document named `output.docx`, place `input.pdf` in the `pdf` directory and run the script. The output will be saved as `output.docx` in the `output` directory.

## Logging

The script uses Python's logging module to provide detailed debug information. Logs include information about the conversion process, OCR performance, and any errors encountered.

## Error Handling

The script includes comprehensive error handling for various scenarios:
- Missing input PDF file
- Poppler not installed or not in PATH
- OCR failures
- Unexpected errors

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request.

## Acknowledgements

- [pdf2image](https://github.com/Belval/pdf2image)
- [pytesseract](https://github.com/madmaze/pytesseract)
- [python-docx](https://github.com/python-openxml/python-docx)
- [Pillow](https://python-pillow.org/)
- [OpenCV](https://opencv.org/)
- [NumPy](https://numpy.org/)