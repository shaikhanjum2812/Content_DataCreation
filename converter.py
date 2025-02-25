import logging
from docx import Document
from openpyxl import Workbook
import pandas as pd
import os
import tempfile
import shutil
from zipfile import ZipFile

def extract_reader_data(doc):
    """Extract data using Reader mode"""
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
                options, answer_part = options_answer.split("answer:")

                hint = ""
                if "Hint:" in answer_part:
                    answer, hint = answer_part.split("Hint:")
                else:
                    answer = answer_part

                answer = answer.strip()
                hint = hint.strip()
                options = options.strip()

                answer_type = "checkbox" if ',' in answer else "radio"
                if answer_type == "radio":
                    try:
                        answer = int(answer)
                    except ValueError:
                        answer = answer

                sheet2_data.append([
                    current_exid, question_key, question.strip(),
                    answer_type, options, answer, hint
                ])
                question_key += 1

            except Exception as e:
                logging.error(f"Error parsing question: {text}. Error: {str(e)}")

    return sheet1_data, sheet2_data

def extract_debug_data(doc):
    """Extract data using Debug mode"""
    ex_data = []
    qa_data = []

    current_entry = {}
    current_exid = ""
    key = 1
    collecting_description = False

    for para in doc.paragraphs:
        text = para.text.strip()

        if text.startswith("exid :"):
            if current_entry:
                ex_data.append(current_entry)
            current_entry = {'exid': text.split("exid :")[1].strip()}
            current_exid = current_entry['exid']
            collecting_description = False
            key = 1
        elif text.startswith("title :"):
            current_entry['title'] = text.split("title :")[1].strip()
            collecting_description = False
        elif text.startswith("description :"):
            current_entry['description'] = text.split("description :")[1].strip()
            collecting_description = True
        elif "assert" in text:
            qa_data.append([current_exid, key, current_entry.get('title', ''), 'assert', '', text])
            key += 1
        elif collecting_description:
            current_entry['description'] += "\n" + text

    if current_entry:
        ex_data.append(current_entry)

    return list(map(lambda x: [
        x.get('exid', ''), x.get('title', ''), x.get('description', ''),
        '', '', 0, '', '', '', 0, 0, 0, '', ''
    ], ex_data)), qa_data

def extract_solver_data(doc):
    """Extract data using Solver mode"""
    ex_data = []
    qa_data = []

    current_entry = {}
    current_exid = ""
    key = 1
    label = ""
    collecting_description = False

    for para in doc.paragraphs:
        text = para.text.strip()

        if text.startswith("exid :"):
            if current_entry:
                ex_data.append(current_entry)
            current_entry = {'exid': text.split("exid :")[1].strip()}
            current_exid = current_entry['exid']
            collecting_description = False
            key = 1
        elif text.startswith("description :"):
            current_entry['description'] = text.split("description :")[1].strip()
            collecting_description = True
            if "Function" in text:
                label_start = text.find("Function") + len("Function")
                label = text[label_start:].strip().split()[0].rstrip('()')
            elif "function" in text:
                label_start = text.find("function") + len("function")
                label = text[label_start:].strip().split()[0].rstrip('()')
        elif "assert" in text:
            qa_data.append([current_exid, key, label, 'assert', '', text])
            key += 1
        elif collecting_description:
            current_entry['description'] += "\n" + text

    if current_entry:
        ex_data.append(current_entry)

    return list(map(lambda x: [
        x.get('exid', ''), x.get('title', ''), x.get('description', ''),
        '', '', 0, '', '', '', 0, 0, 0, '', ''
    ], ex_data)), qa_data

def extract_code_with_indentation(doc):
    """
    Extracts code blocks from a Word document and maintains the original indentation.

    Args:
        doc: Document - The loaded Word document

    Returns:
        queries: list of dicts - A list of dictionaries containing the filename and extracted code.
    """
    queries = []
    collecting_code = False
    current_code = []
    current_qlocation = ""

    for paragraph in doc.paragraphs:
        text = paragraph.text

        if text.strip().startswith("qlocation :"):  # Get the qlocation for file naming
            # Extract only the .txt portion from the qlocation
            qlocation_text = text.split("qlocation :")[1].strip()
            if '.txt' in qlocation_text:
                current_qlocation = qlocation_text.split(',')[0].strip()  # Get text before comma if exists
            else:
                current_qlocation = f"code_{len(queries) + 1}.txt"  # Fallback name if no .txt found
        elif text.strip().startswith("Code:"):  # Start of a code block
            collecting_code = True
            current_code = []  # Reset current code collection
        elif "Answer the following questions:" in text.strip() and collecting_code:
            # End of a code block
            if current_qlocation:
                # Join the code lines preserving original whitespace
                code_content = '\n'.join(current_code)
                queries.append({
                    "qlocation": current_qlocation,
                    "code": code_content
                })
                current_qlocation = ""  # Reset qlocation
            collecting_code = False  # Reset flag
            current_code = []  # Reset collection for the next code block
        elif collecting_code:
            # Add the line with its original indentation
            current_code.append(text)  # Store the raw text with indentation

    # Add the last code block if exists
    if collecting_code and current_code and current_qlocation:
        code_content = '\n'.join(current_code)
        queries.append({
            "qlocation": current_qlocation,
            "code": code_content
        })

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

def convert_word_to_excel(input_path, output_path, mode='reader'):
    """
    Convert a Word document to Excel format with specific formatting based on mode.

    Args:
        input_path (str): Path to input Word document
        output_path (str): Path where Excel file will be saved
        mode (str): Conversion mode ('reader', 'debug', or 'solver')
    """
    try:
        # Load the Word document
        doc = Document(input_path)

        # Extract data based on mode
        if mode == 'debug':
            sheet1_data, sheet2_data = extract_debug_data(doc)
        elif mode == 'solver':
            sheet1_data, sheet2_data = extract_solver_data(doc)
        else:  # reader mode
            sheet1_data, sheet2_data = extract_reader_data(doc)

        # Create DataFrames for both sheets
        columns_sheet1 = ['exid', 'title', 'description', 'category', 
                         'subcategoryid', 'level', 'language', 'qlocation', 
                         'module', 'ex_seq', 'cat_seq', 'subcat_seq', 
                         'league', 'labels']
        columns_sheet2 = ['exid', 'key', 'label', 'type', 'options', 
                         'answer', 'Hint']

        df_sheet1 = pd.DataFrame(sheet1_data, columns=columns_sheet1)
        df_sheet2 = pd.DataFrame(sheet2_data, columns=columns_sheet2)

        # Write to Excel file
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_sheet1.to_excel(writer, sheet_name='ex_data', index=False)
            df_sheet2.to_excel(writer, sheet_name='qa_data', index=False)

        # Create zip file of code extracts
        zip_file_path = create_text_files(input_path)
        logging.debug(f"Created zip file of code extracts: {zip_file_path}")

        logging.debug(f"Successfully converted {input_path} to Excel using {mode} mode")
        return True

    except Exception as e:
        logging.error(f"Error converting file: {str(e)}")
        raise