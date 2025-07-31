# convert_doc.py
import sys
import win32com.client
import pythoncom

def convert(doc_path):
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(doc_path)
    output_path = doc_path.replace(".doc", ".docx")
    doc.SaveAs(output_path, FileFormat=16)
    doc.Close()
    word.Quit()
    return output_path

if __name__ == "__main__":
    doc_path = sys.argv[1]
    output_path = convert(doc_path)
    print(output_path)
