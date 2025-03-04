import logging
from docx import Document
from openpyxl import Workbook
import pandas as pd
import os
import tempfile
import shutil
from zipfile import ZipFile

# Set up logging with more detailed format
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Update the create_test_document function to generate a more comprehensive test
def create_test_document():
    """Creates a test document for conversion testing"""
    doc = Document()

    # Add exercise metadata
    doc.add_paragraph("exid : TEST001")
    doc.add_paragraph("title : Test Exercise")
    doc.add_paragraph("description : This is a test exercise")
    doc.add_paragraph("category : Testing")
    doc.add_paragraph("subcategoryid : TEST")
    doc.add_paragraph("level : 1")
    doc.add_paragraph("language : python")
    doc.add_paragraph("qlocation : test_code.txt")
    doc.add_paragraph("module : test")
    doc.add_paragraph("ex_seq : 1")
    doc.add_paragraph("cat_seq : 1")
    doc.add_paragraph("subcat_seq : 1")
    doc.add_paragraph("league : beginner")
    doc.add_paragraph("labels : test,example")

    # Add questions with different formats
    doc.add_paragraph("Question 1: What is the output of print('Hello, World!')?")
    doc.add_paragraph("Options: Hello World,Hi World,Hello, World!,World Hello answer: 3 Hint: Look at the quotes")

    doc.add_paragraph("Question 2: Select all valid Python data types.")
    doc.add_paragraph("Options: int,float,str,bool answer: 1,2,3,4 Hint: All basic types")

    doc.add_paragraph("Question 3: Calculate the result of 2 + 2")
    doc.add_paragraph("Answer: 4 Hint: Basic arithmetic")

    doc.add_paragraph("Question 4: What are functions in programming?")
    doc.add_paragraph("Answer: Reusable blocks of code Hint: Think about code organization")

    # Add code block
    doc.add_paragraph("C Code:")
    doc.add_paragraph("def test_function():")
    doc.add_paragraph("    print('Hello, World!')")
    doc.add_paragraph("Answer the following questions:")

    # Save test document
    test_path = "test_document.docx"
    doc.save(test_path)
    logging.info(f"Created test document with multiple question types at {test_path}")
    return test_path

def extract_data(doc):
    """Extract data from Word document into two sheets format"""
    sheet1_data = []
    sheet2_data = []

    current_data = {
        'exid': '', 'title': '', 'description': '', 'category': '',
        'subcategoryid': '', 'level': 0, 'language': '', 'qlocation': '',
        'module': '', 'ex_seq': 0, 'cat_seq': 0, 'subcat_seq': 0,
        'league': '', 'labels': ''
    }

    # Process Sheet1 data first
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

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
                logging.debug(f"Extracted {field}: {value}")

                if field == 'labels':  # Last field, append the record
                    sheet1_data.append(list(current_data.values()))
                    logging.info(f"Added exercise data: {current_data['exid']}")
                    current_data = {k: '' if isinstance(v, str) else 0
                                    for k, v in current_data.items()}
                break

    # Process Sheet2 data with improved question handling
    exid = ""
    question_key = 1
    current_question = []
    collecting_question = False

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        if text.startswith("exid :"):
            exid = text.split("exid :")[1].strip()
            question_key = 1
            current_question = []
            collecting_question = False
            logging.debug(f"Processing questions for exid: {exid}")

        elif text.startswith("Answer the following questions:"):
            if current_question:  # If we were collecting a question, process it before resetting
                collecting_question = False
                current_question = []
            continue

        elif text.lower().startswith("question"):
            collecting_question = True
            current_question = [text]
            logging.debug(f"Started collecting question: {text}")

        elif collecting_question and ("Options:" in text or "Answer:" in text):
            # Join collected question lines
            question_text = ' '.join(current_question).strip()

            # Remove "Question X:" prefix if present
            if question_text.lower().startswith("question"):
                try:
                    # Handle both "Question X:" and "Question:" formats
                    parts = question_text.split(":", 1)
                    if len(parts) > 1:
                        question_text = parts[1].strip()
                except Exception:
                    # If splitting fails, keep original text
                    pass

            if "Options:" in text:
                try:
                    options_answer = text.split("Options:")[1]
                    options, answer_hint = options_answer.split("answer:")
                    options = options.strip()

                    # Extract answer and hint
                    if "Hint:" in answer_hint:
                        answer, hint = answer_hint.split("Hint:")
                    else:
                        answer = answer_hint.strip()
                        hint = ""

                    answer = answer.strip()
                    hint = hint.strip()

                    # Determine answer type
                    if ',' in answer:
                        answer_type = "checkbox"
                    else:
                        answer_type = "radio"
                        try:
                            answer = int(answer)
                        except ValueError:
                            pass

                    logging.debug(
                        f"Extracted question {question_key}:\n"
                        f"  Label: {question_text}\n"
                        f"  Type: {answer_type}\n"
                        f"  Options: {options}\n"
                        f"  Answer: {answer}\n"
                        f"  Hint: {hint}"
                    )

                    sheet2_data.append([
                        exid, question_key, question_text,
                        answer_type, options, answer, hint
                    ])
                    question_key += 1

                except Exception as e:
                    logging.error(f"Error parsing question with options: {text}. Error: {str(e)}")

            else:  # Answer: format
                try:
                    answer_hint = text.split("Answer:")[1]

                    # Extract answer and hint
                    if "Hint:" in answer_hint:
                        answer, hint = answer_hint.split("Hint:")
                    else:
                        answer = answer_hint.strip()
                        hint = ""

                    answer = answer.strip()
                    hint = hint.strip()

                    # Determine answer type
                    try:
                        answer = float(answer)
                        answer_type = "number"
                        if answer.is_integer():
                            answer = int(answer)
                    except ValueError:
                        answer_type = "text"

                    logging.debug(
                        f"Extracted question {question_key}:\n"
                        f"  Label: {question_text}\n"
                        f"  Type: {answer_type}\n"
                        f"  Answer: {answer}\n"
                        f"  Hint: {hint}"
                    )

                    sheet2_data.append([
                        exid, question_key, question_text,
                        answer_type, "", answer, hint
                    ])
                    question_key += 1

                except Exception as e:
                    logging.error(f"Error parsing question with direct answer: {text}. Error: {str(e)}")

            # Reset question collection
            current_question = []
            collecting_question = False

        elif collecting_question:
            current_question.append(text)

    return sheet1_data, sheet2_data

def extract_code_with_indentation(doc):
    """
    Extracts code blocks from a Word document and maintains the original indentation.
    The function looks for 'qlocation :' to determine the output filename, falling back to
    CFSF{number}.txt if no qlocation is specified.
    """
    queries = []
    collecting_code = False
    current_code = []
    current_qlocation = ""
    query_count = 1

    for paragraph in doc.paragraphs:
        text = paragraph.text

        # Handle qlocation specification
        if text.strip().startswith("qlocation :"):
            qlocation_text = text.split("qlocation :")[1].strip()
            if '.txt' in qlocation_text:
                current_qlocation = qlocation_text.split(',')[0].strip()
                logging.debug(f"Found qlocation specification: {current_qlocation}")

        elif text.strip().startswith("C Code:") or text.strip().startswith("Code:"):
            collecting_code = True
            current_code = []
            if not current_qlocation:
                current_qlocation = f"CFSF{query_count}.txt"
                logging.debug(f"Using default qlocation: {current_qlocation}")

        elif "Answer the following questions:" in text.strip() and collecting_code:
            if current_code:  # Only add if we have collected code
                code_content = '\n'.join(current_code)
                queries.append({
                    "qlocation": current_qlocation,
                    "code": code_content
                })
                logging.info(f"Extracted code block saved with qlocation: {current_qlocation}")

            current_qlocation = ""
            collecting_code = False
            current_code = []
            query_count += 1

        elif collecting_code:
            current_code.append(text)

    return queries

def create_text_files(input_path):
    """Creates text files from code blocks in a Word document and returns them as a zip file."""
    try:
        doc = Document(input_path)
        temp_dir = tempfile.mkdtemp()
        queries = extract_code_with_indentation(doc)

        for query in queries:
            file_path = os.path.join(temp_dir, query["qlocation"])
            with open(file_path, 'w', encoding='utf-8') as file:
                file.write(query["code"])

        zip_path = os.path.join(tempfile.gettempdir(), 'code_files.zip')
        with ZipFile(zip_path, 'w') as zipf:
            for query in queries:
                file_path = os.path.join(temp_dir, query["qlocation"])
                zipf.write(file_path, query["qlocation"])

        shutil.rmtree(temp_dir)
        return zip_path

    except Exception as e:
        logging.error(f"Error creating text files: {str(e)}")
        raise

def convert_word_to_excel(input_path, output_path):
    """Convert a Word document to Excel format with specific sheets."""
    try:
        doc = Document(input_path)
        sheet1_data, sheet2_data = extract_data(doc)

        # Create DataFrames for both sheets
        columns_sheet1 = ['exid', 'title', 'description', 'category',
                         'subcategoryid', 'level', 'language', 'qlocation',
                         'module', 'ex_seq', 'cat_seq', 'subcat_seq',
                         'league', 'labels']
        columns_sheet2 = ['exid', 'key', 'label', 'type', 'options',
                          'answer', 'hint']

        df_sheet1 = pd.DataFrame(sheet1_data, columns=columns_sheet1)
        df_sheet2 = pd.DataFrame(sheet2_data, columns=columns_sheet2)

        # Write to Excel file
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_sheet1.to_excel(writer, sheet_name='ex_data', index=False)
            df_sheet2.to_excel(writer, sheet_name='qa_data', index=False)

        # Create zip file of code extracts
        zip_file_path = create_text_files(input_path)
        logging.info(f"Created zip file of code extracts: {zip_file_path}")
        logging.info(f"Successfully converted {input_path} to Excel")
        return True

    except Exception as e:
        logging.error(f"Error converting file: {str(e)}")
        raise

# Update the verify_excel_output function to check column names and content
def verify_excel_output(output_path):
    """Verify the structure and content of the generated Excel file"""
    try:
        df_ex = pd.read_excel(output_path, sheet_name='ex_data')
        df_qa = pd.read_excel(output_path, sheet_name='qa_data')

        logging.info("Excel Verification Results:")
        logging.info(f"ex_data columns: {df_ex.columns.tolist()}")
        logging.info(f"qa_data columns: {df_qa.columns.tolist()}")
        logging.info(f"Number of exercises: {len(df_ex)}")
        logging.info(f"Number of questions: {len(df_qa)}")

        # Verify required columns exist
        required_columns = ['exid', 'key', 'label', 'type', 'options', 'answer', 'hint']
        missing_columns = [col for col in required_columns if col not in df_qa.columns]
        if missing_columns:
            logging.error(f"Missing required columns in qa_data: {missing_columns}")
            return False

        # Verify data types
        if not df_qa['key'].dtype.kind in 'ui':  # Check if key is integer
            logging.error("'key' column is not integer type")
            return False

        # Check for empty required fields
        for col in ['exid', 'label', 'type']:
            if df_qa[col].isna().any():
                logging.error(f"Found empty values in required column: {col}")
                return False

        # Verify answer types are valid
        valid_types = {'radio', 'checkbox', 'number', 'text'}
        invalid_types = set(df_qa['type'].unique()) - valid_types
        if invalid_types:
            logging.error(f"Found invalid answer types: {invalid_types}")
            return False

        logging.info("Excel file structure and content verified successfully")
        return True
    except Exception as e:
        logging.error(f"Excel verification failed: {str(e)}")
        return False

if __name__ == "__main__":
    # Create test directories
    if not os.path.exists("output"):
        os.makedirs("output")

    try:
        # Create test document
        test_doc_path = create_test_document()
        logging.info(f"Created test document at {test_doc_path}")

        # Convert to Excel
        output_excel = "output/test_output.xlsx"
        success = convert_word_to_excel(test_doc_path, output_excel)

        if success:
            logging.info(f"Successfully converted document to {output_excel}")
            # Add verification step
            verify_excel_output(output_excel)
            # Clean up test file
            os.remove(test_doc_path)
            logging.info("Test completed successfully")
        else:
            logging.error("Conversion failed")

    except Exception as e:
        logging.error(f"Test failed with error: {str(e)}")