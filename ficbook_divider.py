import os
import tkinter as tk
from tkinter import filedialog
from docx import Document


def open_file():
    ready['text'] = ''
    selected_file['text'] = ''
    file = filedialog.askopenfile(
        mode='r', filetypes=[('Документ Word', '*.docx')])
    if file:
        global filepath
        global name
        global directory
        filepath = os.path.abspath(file.name)
        directory = os.path.dirname(filepath)
        filename = os.path.basename(file.name)
        name, ext = os.path.splitext(os.path.basename(filepath))
        selected_file['text'] = filename


def submit():
    divider_text = divider.get().rstrip('\n')
    if not divider_text:
        ready['text'] = 'Введите разделитель!'
    else:
        try:
            document = Document(filepath)
            segments = []
            current_segment = []
            filename_folder = os.path.join(directory, name)
            os.makedirs(filename_folder, exist_ok=True)

            for para in document.paragraphs:
                if divider_text in para.text:
                    if current_segment:
                        segments.append(current_segment)
                        current_segment = []
                else:
                    current_segment.append(para.text)
            if current_segment:
                segments.append(current_segment)

            for i, segment in enumerate(segments):
                new_doc = Document()
                for text in segment:
                    new_doc.add_paragraph(text)
                    new_doc.save(f'{filename_folder}/split_part_{i+1}.docx')
            ready['text'] = 'Готово!'
        except NameError:
            ready['text'] = 'Укажите файл!'


root = tk.Tk()
root.title('Разделение фанфика на главы')
root.geometry('400x275')

import_button = tk.Button(root, text='Импорт .docx', command=open_file)
import_button.pack(padx=6, pady=6)

selected = tk.Label(text='Выбранный файл:')
selected.pack(padx=3, pady=3)
selected_file = tk.Label(root)
selected_file.pack(padx=6, pady=6)

label = tk.Label(root, text='Разделитель главы:')
label.pack(padx=6, pady=6)
divider = tk.Entry(root)
divider.pack(padx=6, pady=3)

submit_button = tk.Button(root, text='Применить', command=submit)
submit_button.pack(padx=6, pady=10)

ready = tk.Label(root, wraplength=350)
ready.pack(padx=6, pady=6)

root.mainloop()
