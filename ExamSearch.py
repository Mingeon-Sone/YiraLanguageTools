import os

import fitz

from docx import Document



def search_and_read_pdf(file_path, search_word):

    try:

        with fitz.open(file_path) as pdf_doc:

            for page_num in range(pdf_doc.page_count):

                page = pdf_doc[page_num]

                text = page.get_text()

                if search_word.lower() in text.lower():

                    return True

        return False

    except Exception as e:

        print(f"Error reading {file_path} (pdf): {e}")

        return False



def search_and_read_docx(file_path, search_word):

    try:

        doc = Document(file_path)

        for paragraph in doc.paragraphs:

            if search_word.lower() in paragraph.text.lower():

                return True

        return False

    except Exception as e:

        print(f"Error reading {file_path}: {e}")

        return False



def search_in_file(file_path, search_word):

    if file_path.endswith(".docx"):

        return search_and_read_docx(file_path, search_word)

    elif file_path.endswith(".pdf"):

        return search_and_read_pdf(file_path, search_word)

    else:

        return False



if __name__ == "__main__":

    folder_path = r'C:\SynologyDrive\모의고사\모의고사_모음(2010~2023)'



    while True:

        search_word = input("Enter the word to search (나가려면 'exit'를 입력하세요): ")



        if search_word.lower() == 'exit':

            break



        found = False



        for root, dirs, files in os.walk(folder_path):

            for file in files:

                if file.endswith((".docx", ".pdf")) and not file.startswith("~$"):

                    file_path = os.path.join(root, file)



                    if search_in_file(file_path, search_word):

                        print(f"Found '{search_word}' in: {file_path}")

                        found = True



        if not found:

            print(f"'{search_word}' not found in any document.")

