import logging
from docx import Document
from openpyxl import Workbook
import pandas as pd
import os
import tempfile
import shutil
from zipfile import ZipFile

def extract_data(doc):
    """Extract data from Word document into two sheets format"""
    sheet1_data = []
    sheet2_data = []

    current_exid = ""
    current_data = {
        'exid': '', 'title': '', 'description': '', 'category': '',
        'subcategoryid': '', 'level': 0, 'language': '', 'qlocation': '',
        'module': '', 'ex_seq': 0, 'cat_seq': 0, 'subcat_seq': 0,
        'league': '', 'labels': ''
    }
    question_key = 1

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # Process Sheet1 data
        for field in current_data.keys():
            prefix = f"{field} :"
            if text.startswith(prefix):
                value = text.split(prefix)[1].strip()
                if field in ['level', 'ex_seq', 'cat_seq', 'subcat_seq']:
                    try:
                        value = int(value)
                    except ValueError:
                        value = 0
                current_data[field] = value

                if field == 'labels':  # Last field, append the record
                    sheet1_data.append(list(current_data.values()))
                    current_data = {k: '' if isinstance(v, str) else 0 
                                  for k, v in current_data.items()}
                break

        # Process Sheet2 data
        if text.startswith("exid :"):
            current_exid = text.split("exid :")[1].strip()
            question_key = 1
        elif "Options:" in text and "answer:" in text:
            try:
                question, options_answer = text.split("Options:")
                options, answer = options_answer.split("answer:")
                answer = answer.strip()
                options = options.strip()

                # Determine the answer type based on content
                if ',' in answer:
                    answer_type = "checkbox"  # Multiple answers
                elif options:  # Has options but single answer
                    answer_type = "radio"
                else:  # No options
                    try:
                        float(answer)  # Try converting to number
                        answer_type = "number"
                    except ValueError:
                        answer_type = "text"

                sheet2_data.append([
                    current_exid, question_key, question.strip(),
                    answer_type, options, answer
                ])
                question_key += 1

            except Exception as e:
                logging.error(f"Error parsing question: {text}. Error: {str(e)}")

    return sheet1_data, sheet2_data

def extract_code_with_indentation(doc):
    """
    Extracts code blocks from a Word document and maintains the original indentation.
    The function looks for 'qlocation :' to determine the output filename, falling back to
    CFSF{number}.txt if no qlocation is specified.

    Args:
        doc: Document - The loaded Word document

    Returns:
        queries: list of dicts - A list of dictionaries containing the filename and extracted code.
    """
    queries = []
    collecting_code = False
    current_code = []
    current_qlocation = ""
    query_count = 1

    for paragraph in doc.paragraphs:
        text = paragraph.text

        # Handle qlocation specification
        # Example format in Word doc: "qlocation : filename.txt"
        if text.strip().startswith("qlocation :"):
            qlocation_text = text.split("qlocation :")[1].strip()
            if '.txt' in qlocation_text:
                # Extract filename before any comma if present
                # Example: "example.txt, other info" becomes "example.txt"
                current_qlocation = qlocation_text.split(',')[0].strip()
                logging.debug(f"Found qlocation specification: {current_qlocation}")

        # Start collecting code when "C Code:" or "Code:" is found
        elif text.strip().startswith("C Code:") or text.strip().startswith("Code:"):
            collecting_code = True
            current_code = []  # Reset current code collection
            if not current_qlocation:  # If no qlocation was specified before code block
                current_qlocation = f"CFSF{query_count}.txt"
                logging.debug(f"Using default qlocation: {current_qlocation}")

        # Stop collecting code when answer section is reached
        elif "Answer the following questions:" in text.strip() and collecting_code:
            if not current_qlocation:  # Final fallback if somehow no qlocation was set
                current_qlocation = f"CFSF{query_count}.txt"
                logging.debug(f"Using fallback qlocation: {current_qlocation}")

            # Preserve original indentation by joining lines
            code_content = '\n'.join(current_code)
            queries.append({
                "qlocation": current_qlocation,
                "code": code_content
            })
            logging.info(f"Extracted code block saved with qlocation: {current_qlocation}")

            # Reset for next code block
            current_qlocation = ""
            collecting_code = False
            current_code = []
            query_count += 1

        # Collect code lines with original indentation
        elif collecting_code:
            current_code.append(text)

    # Handle the last code block if exists
    if collecting_code and current_code:
        if not current_qlocation:
            current_qlocation = f"CFSF{query_count}.txt"
            logging.debug(f"Using final fallback qlocation: {current_qlocation}")

        code_content = '\n'.join(current_code)
        queries.append({
            "qlocation": current_qlocation,
            "code": code_content
        })
        logging.info(f"Extracted final code block saved with qlocation: {current_qlocation}")

    return queries

def create_text_files(input_path):
    """
    Creates text files from code blocks in a Word document and returns them as a zip file.

    Args:
        input_path (str): Path to input Word document

    Returns:
        str: Path to the created zip file containing text files
    """
    try:
        # Load the Word document
        doc = Document(input_path)

        # Create a temporary directory for the text files
        temp_dir = tempfile.mkdtemp()

        # Extract code blocks and create text files
        queries = extract_code_with_indentation(doc)

        # Create text files in the temporary directory
        for query in queries:
            file_path = os.path.join(temp_dir, query["qlocation"])
            with open(file_path, 'w', encoding='utf-8') as file:
                file.write(query["code"])

        # Create a zip file containing all text files
        zip_path = os.path.join(tempfile.gettempdir(), 'code_files.zip')
        with ZipFile(zip_path, 'w') as zipf:
            for query in queries:
                file_path = os.path.join(temp_dir, query["qlocation"])
                zipf.write(file_path, query["qlocation"])

        # Clean up the temporary directory
        shutil.rmtree(temp_dir)

        return zip_path

    except Exception as e:
        logging.error(f"Error creating text files: {str(e)}")
        raise

def convert_word_to_excel(input_path, output_path):
    """
    Convert a Word document to Excel format with specific sheets.

    Args:
        input_path (str): Path to input Word document
        output_path (str): Path where Excel file will be saved
    """
    try:
        # Load the Word document
        doc = Document(input_path)

        # Extract data
        sheet1_data, sheet2_data = extract_data(doc)

        # Create DataFrames for both sheets
        columns_sheet1 = ['exid', 'title', 'description', 'category', 
                         'subcategoryid', 'level', 'language', 'qlocation', 
                         'module', 'ex_seq', 'cat_seq', 'subcat_seq', 
                         'league', 'labels']
        columns_sheet2 = ['exid', 'key', 'question', 'type', 'options', 
                         'answer']

        df_sheet1 = pd.DataFrame(sheet1_data, columns=columns_sheet1)
        df_sheet2 = pd.DataFrame(sheet2_data, columns=columns_sheet2)

        # Write to Excel file
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_sheet1.to_excel(writer, sheet_name='ex_data', index=False)
            df_sheet2.to_excel(writer, sheet_name='qa_data', index=False)

        # Create zip file of code extracts
        zip_file_path = create_text_files(input_path)
        logging.debug(f"Created zip file of code extracts: {zip_file_path}")

        logging.debug(f"Successfully converted {input_path} to Excel")
        return True

    except Exception as e:
        logging.error(f"Error converting file: {str(e)}")
        raise