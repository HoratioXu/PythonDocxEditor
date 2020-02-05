from win32com import client as wc

PATH = "C:\\Users\horatio xu\\Desktop\\cover_letter.docx"
OUTPUT = "C:\\Users\horatio xu\\Desktop\\cover_letter.pdf"

# for multiple conversions
# def get_docx(input_docx):
#     docx_path = []
#     for root,dirs,filenames in os.walk(input_docx):
#         for filename in filenames:
#             if filename.endswith(('.docx','.doc')):
#                 docx_path.append(root+'/'+filename)
#     return docx_path


def docx_to_pdf(docx_path, pdf_path):
    print('docx_path', docx_path)
    word = wc.Dispatch('Word.Application')
    word.Visible = 0

    try:
        doc = word.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, 17)
        doc.Close()
        print('%s PDF Success' % docx_path)
    except:
        print('%s PDF Failed' % docx_path)
    word.Quit()


if __name__ == '__main__':
    docx_to_pdf(PATH, OUTPUT)
