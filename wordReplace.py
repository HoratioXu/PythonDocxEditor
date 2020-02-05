from docx import Document
import os

OLDPATH = "C:\\Users\\horatio xu\\Desktop\\files\\personal\\CV"
PATH = "C:\\Users\\horatio xu\\Desktop"
DICT = {
    "send_date": "Feb 05 2020",
    "namecompany": "MicroSoft",
    "company_address": "Vancouver, BC, CA"
}

def main():
    for fileName in os.listdir(OLDPATH):
        oldFile = OLDPATH + "\\" + fileName
        newFile = PATH + "\\" + fileName
        if oldFile.split(".")[1] == 'docx':
            document = Document(oldFile)
            document = check(document)
            document.save(newFile)


def check(document):
    # tables
    for table in document.tables:
        for row in range(len(table.rows)):
            for col in range(len(table.columns)):
                for key, value in DICT.items():
                    if key in table.cell(row, col).text:
                        print(key + "->" + value)
                        table.cell(row, col).text = table.cell(row, col).text.replace(key, value)

    # paragraphs
    for para in document.paragraphs:
        for i in range(len(para.runs)):
            for key, value in DICT.items():
                if key in para.runs[i].text:
                    print(key + "->" + value)
                    para.runs[i].text = para.runs[i].text.replace(key, value)

    return document


if __name__ == '__main__':
    main()