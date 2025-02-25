# Word to Excel Converter

A web application that converts Word documents to Excel format with robust file processing capabilities.

## Features

- Convert Word documents (.docx) to Excel format
- Multiple conversion modes:
  - Reader Mode: Extracts structured data including exercise IDs, titles, descriptions, and question details
  - Debug Mode: Extracts debugging information and assertion statements
  - Solver Mode: Extracts solver-specific data and function implementations
- Generates text files from code blocks in Word documents
- Clean and user-friendly interface
- Error handling and validation

## Tech Stack

- Flask web framework
- Python-docx for document handling
- Openpyxl for Excel conversion
- Pandas for data manipulation
- Bootstrap for UI

## Setup and Installation

1. Clone the repository:
```bash
git clone <your-repository-url>
cd word-to-excel-converter
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
python main.py
```

The application will be available at `http://localhost:5000`

## Usage

1. Visit the application in your web browser
2. Upload a Word document (.docx)
3. Select the conversion mode (Reader, Debug, or Solver)
4. Click "Convert to Excel" to process and download the converted file
5. Optionally, use "Generate Text Files" to extract code blocks as text files

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License

[MIT](https://choosealicense.com/licenses/mit/)
