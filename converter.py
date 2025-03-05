import os
import logging
from docx import Document
from openpyxl import Workbook
import pandas as pd
import tempfile
import shutil
from zipfile import ZipFile

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

def extract_sheet1_data_from_docx(docx_path):
    """Extract data from the Word document for Sheet1"""
    document = Document(docx_path)
    data = []

    current_exid = ""
    current_title = ""
    current_description = ""
    current_category = ""
    current_subcategoryid = ""
    current_level = 0
    current_language = ""
    current_qlocation = ""
    current_module = ""
    current_ex_seq = 0
    current_cat_seq = 0
    current_subcat_seq = 0
    current_league = ""
    current_labels = ""

    for para in document.paragraphs:
        text = para.text.strip()

        if text.startswith("exid :"):
            if current_exid:
                data.append([
                    current_exid, current_title, current_description, current_category,
                    current_subcategoryid, current_level, current_language, current_qlocation,
                    current_module, current_ex_seq, current_cat_seq, current_subcat_seq,
                    current_league, current_labels
                ])
            current_exid = text.split("exid :")[1].strip()
        elif text.startswith("title :"):
            current_title = text.split("title :")[1].strip()
        elif text.startswith("description :"):
            current_description = text.split("description :")[1].strip()
        elif text.startswith("category :"):
            current_category = text.split("category :")[1].strip()
        elif text.startswith("subcategoryid :"):
            current_subcategoryid = text.split("subcategoryid :")[1].strip()
        elif text.startswith("level :"):
            try:
                current_level = int(text.split("level :")[1].strip())
            except ValueError:
                current_level = 0
        elif text.startswith("language :"):
            current_language = text.split("language :")[1].strip()
        elif text.startswith("qlocation :"):
            current_qlocation = text.split("qlocation :")[1].strip()
        elif text.startswith("module :"):
            current_module = text.split("module :")[1].strip()
        elif text.startswith("ex_seq :"):
            try:
                current_ex_seq = int(text.split("ex_seq :")[1].strip())
            except ValueError:
                current_ex_seq = 0
        elif text.startswith("cat_seq :"):
            try:
                current_cat_seq = int(text.split("cat_seq :")[1].strip())
            except ValueError:
                current_cat_seq = 0
        elif text.startswith("subcat_seq :"):
            try:
                current_subcat_seq = int(text.split("subcat_seq :")[1].strip())
            except ValueError:
                current_subcat_seq = 0
        elif text.startswith("league :"):
            current_league = text.split("league :")[1].strip()
        elif text.startswith("labels :"):
            current_labels = text.split("labels :")[1].strip()

    if current_exid:
        data.append([
            current_exid, current_title, current_description, current_category,
            current_subcategoryid, current_level, current_language, current_qlocation,
            current_module, current_ex_seq, current_cat_seq, current_subcat_seq,
            current_league, current_labels
        ])

    return data

def extract_sheet2_data_from_docx(docx_path):
    """Extract questions and answers from the Word document for Sheet2"""
    document = Document(docx_path)
    questions_data = []
    exid = ""
    question_key = 1

    for para in document.paragraphs:
        text = para.text.strip()

        if text.startswith("exid :"):
            exid = text.split("exid :")[1].strip()
            question_key = 1
        elif text.startswith("Answer the following questions:"):
            continue
        elif "Options:" in text and "answer:" in text:
            try:
                question, options_answer = text.split("Options:")
                options, answer = options_answer.split("answer:")
                options = options.strip().split(',')
                answer = answer.strip()

                if ',' in answer:
                    answer_type = "checkbox"
                else:
                    answer_type = "radio"
                    try:
                        answer = int(answer)
                    except ValueError:
                        pass

                questions_data.append([
                    exid, question_key, question.strip(), 
                    answer_type, ','.join(options), answer
                ])
                question_key += 1
            except ValueError:
                logger.error(f"Warning: Couldn't parse question options and answer in: {text}")
                continue
        elif "Answer:" in text:
            try:
                question, answer = text.split("Answer:")
                answer = answer.strip()
                try:
                    answer = float(answer)
                    answer_type = "number"
                    if answer.is_integer():
                        answer = int(answer)
                except ValueError:
                    answer_type = "text"

                questions_data.append([
                    exid, question_key, question.strip(),
                    answer_type, "", answer
                ])
                question_key += 1
            except ValueError:
                logger.error(f"Warning: Couldn't parse question and answer in: {text}")
                continue

    return questions_data

def convert_word_to_excel(input_path, output_path):
    """Convert Word document to Excel format with exercise and QA data"""
    try:
        # Extract data from the Word document
        sheet1_data = extract_sheet1_data_from_docx(input_path)
        sheet2_data = extract_sheet2_data_from_docx(input_path)

        # Create DataFrames for Sheet1 and Sheet2
        columns_sheet1 = ['exid', 'title', 'description', 'category', 'subcategoryid', 
                         'level', 'language', 'qlocation', 'module', 'ex_seq', 
                         'cat_seq', 'subcat_seq', 'league', 'labels']
        columns_sheet2 = ['exid', 'key', 'question', 'type', 'options', 'answer']

        df_sheet1 = pd.DataFrame(sheet1_data, columns=columns_sheet1)
        df_sheet2 = pd.DataFrame(sheet2_data, columns=columns_sheet2)

        # Write to Excel file
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_sheet1.to_excel(writer, sheet_name='ex_data', index=False)
            df_sheet2.to_excel(writer, sheet_name='qa_data', index=False)

        logger.info(f"Successfully converted {input_path} to Excel")
        logger.info(f"Sheet1 rows: {len(df_sheet1)}")
        logger.info(f"Sheet2 rows: {len(df_sheet2)}")
        return True

    except Exception as e:
        logger.error(f"Error converting file: {str(e)}")
        raise

def create_text_files(input_path):
    """Creates text files from code blocks in a Word document and returns a zip file path."""
    try:
        doc = Document(input_path)
        temp_dir = tempfile.mkdtemp()

        # Extract code blocks
        queries = []
        collecting_code = False
        current_code = []
        qlocation = None

        for para in doc.paragraphs:
            text = para.text.strip()

            if text.startswith("qlocation :"):
                qlocation = text.split("qlocation :")[1].strip()
                if not qlocation.endswith('.txt'):
                    qlocation = f"{qlocation}.txt"

            elif text.startswith("Code:"):
                collecting_code = True
                current_code = []

            elif collecting_code and "Answer the following questions:" in text:
                if current_code and qlocation:
                    queries.append({
                        "qlocation": qlocation,
                        "code": "\n".join(current_code)
                    })
                collecting_code = False
                current_code = []

            elif collecting_code:
                current_code.append(text)

        # Create zip file with text files
        zip_path = os.path.join(tempfile.gettempdir(), 'code_files.zip')
        with ZipFile(zip_path, 'w') as zipf:
            for query in queries:
                file_path = os.path.join(temp_dir, query["qlocation"])
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(query["code"])
                zipf.write(file_path, query["qlocation"])

        shutil.rmtree(temp_dir)
        return zip_path

    except Exception as e:
        logger.error(f"Error creating text files: {str(e)}")
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