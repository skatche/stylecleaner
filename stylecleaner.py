import tkinter as tk
from tkinter import filedialog as fd
from docx import Document

def strip_styles(input_name, output_name):
    doc = Document(input_name)
    
    style_names = [s.name for s in doc.styles]
    # For some reason deleting the 'Normal' style is precisely what makes it
    # appear in the style gallery, where it can't be deleted, only hidden.
    # So we exclude it here.
    bad_names = [n for n in style_names if n[0:2] != 'MS' and n != 'Normal']
    for n in bad_names:
        doc.styles[n].delete()

    doc.save(output_name)
    return

def gui():
    master = tk.Tk()

    tk.Label(master, text='Input').grid(row=0)
    tk.Label(master, text='Output').grid(row=1)

    input_entry = tk.Entry(master)
    output_entry = tk.Entry(master)

    input_entry.grid(row=0, column=1)
    output_entry.grid(row=1, column=1)

    filetypes = [('Docx file', ('*.docx')),
                 ('All files', ('*.*'))]
    def set_input_file():
        filename = fd.askopenfilename(filetypes=filetypes)
        input_entry.delete(0, tk.END)
        input_entry.insert(0, filename)
        return

    def set_output_file():
        filename = fd.asksaveasfilename(filetypes=filetypes,
                                        defaultextension='.docx')
        output_entry.delete(0, tk.END)
        output_entry.insert(0, filename)
        return

    def clean_doc():
        strip_styles(input_entry.get(), output_entry.get())
        return

    input_browse = tk.Button(master, text='Browse', command=set_input_file)
    output_browse = tk.Button(master, text='Browse', command=set_output_file)
    
    input_browse.grid(row=0, column=2)
    output_browse.grid(row=1, column=2)

    clean = tk.Button(master, text='Clean', command=clean_doc)
    clean.grid(row=2, column=2)

    tk.mainloop()
    return

if __name__ == '__main__':
    gui()



