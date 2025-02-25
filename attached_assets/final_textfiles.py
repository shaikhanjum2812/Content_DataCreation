import os
from docx import Document

def extract_queries_from_docx(docx_file):
    """
    Extracts SQL queries from a Word document that start with "SQL Query" and end with "label",
    excluding the starting and ending markers from the text files.
    
    Args:
    docx_file: str - The path to the Word document (.docx) from which to extract queries.
    
    Returns:
    queries: list of dicts - A list of dictionaries containing the filename and extracted query.
    """
    document = Document(docx_file)
    queries = []
    query_count = 1
    collecting_query = False
    current_query = []

    for paragraph in document.paragraphs:
        text = paragraph.text.strip()

        if text.lower().startswith("C Code:"):
            # Start collecting the query but don't include "SQL Query" in the output
            collecting_query = True
            current_query = []  # Initialize an empty list to start collecting

        elif "Answer the following questions:" in text.lower() and collecting_query:
            # End collecting the query when "label" is found (exclude the label)
            qlocation = f"CFSF{query_count}.txt"
            queries.append({
                "qlocation": qlocation,
                "query": "\n".join(current_query)  # Join the query lines
            })
            query_count += 1
            collecting_query = False  # Reset for the next query
            current_query = []  # Clear the current query

        elif collecting_query:
            # Collect lines while the query is being processed
            current_query.append(text)

    return queries

def create_text_files_from_queries(queries, output_dir):
    """
    Create text files from extracted queries and save them in a specified folder.
    
    Args:
    queries: list of dicts - List of extracted queries and file names.
    output_dir: str - The directory where the files will be saved.
    """
    # Check if the directory exists, if not, create it
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for query in queries:
        file_path = os.path.join(output_dir, query["qlocation"])
        with open(file_path, 'w') as file:
            file.write(query["query"])
        print(f"File created: {file_path}")

# Example usage
if __name__ == "__main__":
    # Specify the path to your Word document
    docx_file = "Z:/DATASCIENCELAB/Projects/Pravinyam_Project/CProgramFiles/Functions _ static functions  (1).docx"  # Change this to your actual document path
    
    # Specify the output folder
    output_folder = "Z:/DATASCIENCELAB/Projects/Pravinyam_Project/CProgramFiles/Static_functions"  # Change this to your desired folder name
    
    # Extract SQL queries from the Word document
    queries = extract_queries_from_docx(docx_file)
    
    # Create the text files with the extracted SQL queries in the specified folder
    create_text_files_from_queries(queries, output_folder)
