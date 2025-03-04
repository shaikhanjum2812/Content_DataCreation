import logging
from docx import Document
from openpyxl import Workbook
import pandas as pd
import os
import tempfile
import shutil
from zipfile import ZipFile

# Set up logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def extract_sheet2_data(doc):
    """Extract QA data from Word document"""
    questions_data = []
    current_exid = None
    question_key = 1
    in_qa_section = False

    logger.info("Starting QA data extraction...")

    # First get EXID
    for para in doc.paragraphs:
        text = para.text.strip()
        if text.startswith("exid :"):
            current_exid = text.split("exid :")[1].strip()
            logger.info(f"Found EXID: {current_exid}")
            break

    if not current_exid:
        logger.error("No EXID found in document")
        return []

    # Now process paragraphs
    for i in range(len(doc.paragraphs)):
        text = doc.paragraphs[i].text.strip()

        if not text:
            continue

        logger.debug(f"Processing paragraph {i}: {text}")

        if "Answer the following questions:" in text:
            in_qa_section = True
            logger.info("Entered QA section")
            continue

        if not in_qa_section:
            continue

        if text.lower().startswith("question"):
            try:
                # Extract question text
                question_text = text.split(":", 1)[1].strip() if ":" in text else text
                logger.info(f"Found question: {question_text}")

                # Look ahead for the answer line
                if i + 1 < len(doc.paragraphs):
                    answer_text = doc.paragraphs[i + 1].text.strip()
                    logger.debug(f"Processing answer: {answer_text}")

                    if "Options:" in answer_text and "answer:" in answer_text:
                        # Handle MCQ format
                        options_part = answer_text.split("Options:")[1]
                        options, answer_part = options_part.split("answer:")

                        options = options.strip()
                        answer = answer_part.strip()
                        hint = ""

                        if "Hint:" in answer:
                            answer, hint = answer.split("Hint:")
                            answer = answer.strip()
                            hint = hint.strip()

                        answer_type = "checkbox" if "," in answer else "radio"
                        try:
                            if answer_type == "radio":
                                answer = int(answer)
                        except ValueError:
                            pass

                        qa_entry = [
                            current_exid,
                            question_key,
                            question_text,
                            answer_type,
                            options,
                            answer,
                            hint
                        ]
                        questions_data.append(qa_entry)
                        logger.info(f"Added MCQ - Key: {question_key}, Type: {answer_type}")
                        logger.debug(f"QA Entry: {qa_entry}")
                        question_key += 1

                    elif "Answer:" in answer_text:
                        # Handle direct answer format
                        answer_part = answer_text.split("Answer:")[1]
                        answer = answer_part.strip()
                        hint = ""

                        if "Hint:" in answer:
                            answer, hint = answer.split("Hint:")
                            answer = answer.strip()
                            hint = hint.strip()

                        try:
                            answer = float(answer)
                            answer_type = "number"
                            if answer.is_integer():
                                answer = int(answer)
                        except ValueError:
                            answer_type = "text"

                        qa_entry = [
                            current_exid,
                            question_key,
                            question_text,
                            answer_type,
                            "",
                            answer,
                            hint
                        ]
                        questions_data.append(qa_entry)
                        logger.info(f"Added Direct Answer - Key: {question_key}, Type: {answer_type}")
                        logger.debug(f"QA Entry: {qa_entry}")
                        question_key += 1

            except Exception as e:
                logger.error(f"Error processing QA pair: {str(e)}")
                logger.error(f"Question: {text}")
                if i + 1 < len(doc.paragraphs):
                    logger.error(f"Answer: {doc.paragraphs[i + 1].text.strip()}")

    logger.info(f"Total questions extracted: {len(questions_data)}")
    return questions_data

def extract_sheet1_data(doc):
    """Extract exercise metadata from Word document"""
    exercises_data = []
    current_data = {
        'exid': '', 'title': '', 'description': '', 'category': '',
        'subcategoryid': '', 'level': 0, 'language': '', 'qlocation': '',
        'module': '', 'ex_seq': 0, 'cat_seq': 0, 'subcat_seq': 0,
        'league': '', 'labels': ''
    }

    logger.info("Starting exercise data extraction...")

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
                logger.debug(f"Extracted {field}: {value}")

                if field == 'labels':
                    exercises_data.append(list(current_data.values()))
                    logger.info(f"Added exercise data: {current_data['exid']}")
                    current_data = {k: '' if isinstance(v, str) else 0 
                                  for k, v in current_data.items()}
                break

    logger.info(f"Total exercises extracted: {len(exercises_data)}")
    return exercises_data

def create_test_document():
    """Creates a test document for QA extraction"""
    doc = Document()

    # Add exercise metadata
    doc.add_paragraph("exid : TEST001")
    doc.add_paragraph("title : Test Exercise")
    doc.add_paragraph("description : This is a test exercise")
    doc.add_paragraph("category : Testing")
    doc.add_paragraph("subcategoryid : TEST")
    doc.add_paragraph("level : 1")
    doc.add_paragraph("language : python")
    doc.add_paragraph("qlocation : test.txt")
    doc.add_paragraph("module : test")
    doc.add_paragraph("ex_seq : 1")
    doc.add_paragraph("cat_seq : 1")
    doc.add_paragraph("subcat_seq : 1")
    doc.add_paragraph("league : beginner")
    doc.add_paragraph("labels : test,example")

    # Add empty line before questions
    doc.add_paragraph("")

    # Add questions section marker
    doc.add_paragraph("Answer the following questions:")
    doc.add_paragraph("")

    # Add MCQ with hint
    doc.add_paragraph("Question 1: What is the output of print('Hello, World!')?")
    doc.add_paragraph("Options: Hello World,Hi World,Hello, World!,World Hello answer: 3 Hint: Look at the quotes")

    # Add Checkbox with hint
    doc.add_paragraph("Question 2: What are valid Python data types?")
    doc.add_paragraph("Options: int,float,str,bool answer: 1,2,3,4 Hint: All basic types")

    # Add Number answer with hint
    doc.add_paragraph("Question 3: What is 2 + 2?")
    doc.add_paragraph("Answer: 4 Hint: Basic arithmetic")

    # Add Text answer with hint
    doc.add_paragraph("Question 4: What are functions in programming?")
    doc.add_paragraph("Answer: Reusable blocks of code Hint: Think about code organization")

    test_path = "test_document.docx"
    doc.save(test_path)
    logger.info(f"Created test document: {test_path}")
    return test_path

def convert_word_to_excel(input_path, output_path):
    """Convert Word document to Excel format with exercise and QA data"""
    try:
        logger.info(f"Starting conversion of {input_path}")
        doc = Document(input_path)

        # Extract data
        sheet1_data = extract_sheet1_data(doc)
        sheet2_data = extract_sheet2_data(doc)

        # Log extraction results
        logger.info("\nExtraction Summary:")
        logger.info(f"Number of exercises: {len(sheet1_data)}")
        logger.info(f"Number of questions: {len(sheet2_data)}")

        if not sheet1_data or not sheet2_data:
            logger.error("No data extracted!")
            return False

        # Create DataFrames
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

        logger.info(f"Successfully wrote data to Excel: {output_path}")
        logger.info(f"Sheet1 rows: {len(df_sheet1)}")
        logger.info(f"Sheet2 rows: {len(df_sheet2)}")
        return True

    except Exception as e:
        logger.error(f"Error converting file: {str(e)}")
        raise

if __name__ == "__main__":
    # Create output directory if it doesn't exist
    if not os.path.exists("output"):
        os.makedirs("output")

    try:
        # Create and process test document
        test_doc_path = create_test_document()
        output_excel = "output/test_output.xlsx"
        success = convert_word_to_excel(test_doc_path, output_excel)

        if success:
            logger.info(f"Successfully converted document to {output_excel}")
            os.remove(test_doc_path)
            logger.info("Test completed successfully")
        else:
            logger.error("Conversion failed")

    except Exception as e:
        logger.error(f"Test failed with error: {str(e)}")