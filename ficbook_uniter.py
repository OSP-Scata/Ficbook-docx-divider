import os
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
import tkinter.messagebox as messagebox
from docx import Document
from docxcompose.composer import Composer

dropped_files = []


def drop_files(event):  # обработка перетащенных файлов и сохранение путей к ним
    files = root.tk.splitlist(event.data)

    for file_path in files:
        file_path = file_path.strip('{}')
        if file_path.endswith('.docx'):
            if file_path not in dropped_files:
                dropped_files.append(file_path)
                listbox.insert("end", os.path.basename(file_path))
        else:
            messagebox.showwarning(
                "Ошибка", f"Файл {os.path.basename(file_path)} не является .docx")


def clear_list():  # функция для очистки списка файлов
    dropped_files.clear()
    listbox.delete(0, "end")


def merge_docx_files():  # объединение файлов
    merged_file_name = merged_file.get().strip()

    if not dropped_files:
        messagebox.showwarning(
            "Внимание", "Сначала перетащите DOCX-файлы в окно!")
        return

    if len(dropped_files) < 2:
        messagebox.showwarning(
            "Внимание", "Для объединения нужно минимум 2 файла!")
        return

    try:
        master_doc = Document(dropped_files[0])
        composer = Composer(master_doc)

        for file_path in dropped_files[1:]:
            master_doc.add_page_break()
            doc_to_append = Document(file_path)
            composer.append(doc_to_append)

        if not merged_file_name:
            output_path = os.path.join(os.path.dirname(
                dropped_files[0]), "Объединённый документ.docx")
        else:
            output_path = os.path.join(os.path.dirname(
                dropped_files[0]), f"{merged_file_name}.docx")

        composer.save(output_path)
        messagebox.showinfo(
            "Успех", f"Файлы успешно объединены!\nСохранено в:\n{output_path}")

        # Автоматическая очистка после успешного слияния
        clear_list()

    except Exception as e:
        messagebox.showerror("Ошибка обработки",
                             f"Произошла ошибка при сборке: {str(e)}")


# GUI
root = TkinterDnD.Tk()
root.title("Объединение DOCX файлов")
root.geometry("500x530")

drop_zone = tk.Label(root, text="Перетащите сюда DOCX-файлы",
                     bg="#e1f5fe", font=("Arial", 12, "bold"))
drop_zone.pack(fill="x", padx=20, pady=15, ipady=20)

listbox = tk.Listbox(root, height=12, font=("Arial", 10))
listbox.pack(fill="both", expand=True, padx=20, pady=(0, 5))

clear_button = tk.Button(root, text="Очистить список файлов",
                         fg="red", command=clear_list)
clear_button.pack(padx=20, pady=(0, 10))

label = tk.Label(
    root, text='Название финального файла (можно оставить пустым):')
label.pack(padx=6, pady=6)
merged_file = tk.Entry(root)
merged_file.pack(padx=6, pady=3)

merge_button = tk.Button(root, text="Объединить файлы",
                         command=merge_docx_files, bg="#c8e6c9")
merge_button.pack(fill="x", padx=20, pady=(10, 20), ipady=5)

drop_zone.drop_target_register(DND_FILES)
drop_zone.dnd_bind('<<Drop>>', drop_files)

root.mainloop()
