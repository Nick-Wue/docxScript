
import tkinter as tk
from tkinter import filedialog
from docx import Document
import docxedit

if __name__ == '__main__':

    root = tk.Tk()
    root.withdraw()

    path = filedialog.askopenfilename()
    doc = Document(path)
    # adds <lb/> to any linebreak
    docxedit.replace_string(doc, "\n", "<lb/>\n ")

    doc.save("done_" + path)
