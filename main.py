import pyttsx3
import PyPDF2
from docx import Document

def read_pdf(file_path):
    with open(file_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        num_pages = len(pdf_reader.pages)
        text = ""
        for page_num in range(num_pages):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()
        return text

def read_word(file_path):
    doc = Document(file_path)
    text = ""
    for paragraph in doc.paragraphs:
        text += paragraph.text + '\n'
    return text

def speak(text):
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()

def main():
    file_path = input("Enter the path of the document : ")

    if file_path.lower().endswith('.pdf'):
        text = read_pdf(file_path)
    elif file_path.lower().endswith(('.docx', '.doc')):
        text = read_word(file_path)
    else:
        print("Sorry...Unsupported file format!!! Please provide a PDF or Word document.")
        return

    speak(text)

if __name__ == "__main__":
    main()
