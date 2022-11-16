from docx import Document
import docxedit

if __name__ == '__main__':
    doc = Document("binder.weibliche_aerzte.docx")
    # adds <lb/> to any linebreak
    docxedit.replace_string(doc, "\n", "<lb/>\n ")


    doc.save("test.docx")
