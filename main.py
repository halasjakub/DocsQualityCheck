import os
import pyautogui
import tkinter as tk
from tkinter import filedialog, Label, Menu, ttk
from pdf2image import convert_from_path
from PIL import Image, ImageTk

# Lista do przechowywania nazw plików, które zostały przekonwertowane
converted_files = []


def convert_pdf_to_jpg(pdf_file_path):
    """Converts a PDF file to JPG images and stores them in a specific folder."""
    try:
        output_folder = "converted_jpgs"

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            print(f"Folder created: {output_folder}")

        images = convert_from_path(pdf_file_path, dpi=300)
        jpg_file_path = os.path.join(
            output_folder, os.path.basename(pdf_file_path).replace(".pdf", ".jpg")
        )
        images[0].save(jpg_file_path, "JPEG")
        print(f"PDF converted to JPG: {jpg_file_path}")

        # Dodanie nazwy pliku JPG do listy
        converted_files.append(os.path.basename(jpg_file_path))

        # Odświeżenie tabeli po dodaniu pliku
        update_document_number_column()

    except Exception as e:
        print(f"Error converting PDF to JPG: {e}")


def open_and_check_file_format():
    """Opens a file dialog to select multiple files and checks their format (DOCX, XLSX, PDF, TXT)."""
    file_paths = filedialog.askopenfilenames(
        title="Select files", filetypes=[("All Files", "*.*")]
    )

    if not file_paths:
        print("No files selected.")
        return

    for file_path in file_paths:
        file_extension = os.path.splitext(file_path)[1].lower()

        if file_extension == ".docx":
            print(f"DOCX file selected: {file_path}")
        elif file_extension == ".xlsx":
            print(f"XLSX file selected: {file_path}")
        elif file_extension == ".pdf":
            convert_pdf_to_jpg(file_path)
        elif file_extension == ".txt":
            print(f"TXT file selected: {file_path}")
        else:
            print(f"Unsupported file format: {file_extension}")


def empty_function():
    """Test function triggered."""
    print("Test function triggered")


def clear_data():
    """Clear table contents and delete all JPG files in the converted_jpgs folder."""
    # Czyszczenie tabeli
    for row in tree.get_children():
        tree.delete(row)

    # Usuwanie plików JPG z folderu
    output_folder = "converted_jpgs"

    try:
        # Sprawdź, czy folder istnieje
        if os.path.exists(output_folder):
            # Przejrzyj pliki w folderze
            for filename in os.listdir(output_folder):
                # Sprawdź, czy plik ma rozszerzenie .jpg
                if filename.lower().endswith(".jpg"):
                    file_path = os.path.join(output_folder, filename)
                    try:
                        os.remove(file_path)  # Usuń plik JPG
                        print(f"Deleted: {file_path}")
                    except Exception as e:
                        print(f"Error deleting file {file_path}: {e}")
        else:
            print(f"Folder {output_folder} does not exist.")
    except Exception as e:
        print(f"Error clearing data: {e}")


# Główne okno GUI
gui_main_window = tk.Tk()
gui_main_window.title("DocsQualityCheck")
gui_main_window.geometry(f"{pyautogui.size()[0]}x{pyautogui.size()[1]}")


# Pasek menu
menu_bar = Menu(gui_main_window)
gui_main_window.config(menu=menu_bar)

# Menu pliku
file_menu = Menu(menu_bar, tearoff=False)
menu_bar.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Open", command=open_and_check_file_format)
file_menu.add_command(label="Close", command=clear_data)  # Zmieniono na clear_data

# Menu tabeli
table_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Table", menu=table_menu)


def show_basic_table():
    """Display only the 'Documentnumber' column."""
    tree["columns"] = ("documentNumber",)
    for column in tree["columns"]:
        tree.column(column)
        tree.heading(column, text=column.replace("_", " ").title())
    update_document_number_column()


def show_fdf_table():
    """Display 'Documentnumber', 'Study', 'Studysite', 'Investigatorname' columns."""
    tree["columns"] = (
        "documentNumber", "study", "studySite", "investigatorName"
    )
    for column in tree["columns"]:
        tree.column(column)
        tree.heading(column, text=column.replace("_", " ").title())
    update_document_number_column()


table_menu.add_command(label="Basic", command=show_basic_table)
table_menu.add_command(label="FDF", command=show_fdf_table)

# Menu danych
data_menu = Menu(menu_bar, tearoff=False)
menu_bar.add_cascade(label="Data", menu=data_menu)
data_menu.add_command(label="Import", command=empty_function)
data_menu.add_command(label="Export", command=empty_function)

# Menu pomocy
help_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Help", menu=help_menu)
help_menu.add_command(label="About", command=empty_function)
help_menu.add_command(label="Manual", command=empty_function)
help_menu.add_command(label="What's New", command=empty_function)
help_menu.add_command(label="Report", command=empty_function)

# Tabela
tree = ttk.Treeview(gui_main_window, selectmode="browse")
tree.pack(side="left")
verscrlbar = ttk.Scrollbar(gui_main_window, orient="vertical", command=tree.yview)
verscrlbar.pack(side="left", fill="x")
tree.configure(xscrollcommand=verscrlbar.set)

# Ustawienie początkowych kolumn
tree["columns"] = (
    "documentNumber", "study", "studySite", "investigatorName",
    "financialDisclosureInterest", "documentDate", "studySponsor",
    "studyCoSponsor", "signature", "description"
)
tree['show'] = 'headings'

columns = [
    "documentNumber", "study", "studySite", "investigatorName",
    "financialDisclosureInterest", "documentDate", "studySponsor",
    "studyCoSponsor", "signature", "description"
]

# Ustawienie kolumn tabeli
for column in columns:
    tree.column(column)
    tree.heading(column, text=column.replace("_", " ").title())


def update_document_number_column():
    """Update the 'Documentnumber' column with the names of converted files."""
    # Czyszczenie istniejących danych w tabeli
    for row in tree.get_children():
        tree.delete(row)

    # Dodanie nazw plików do tabeli
    for file_name in converted_files:
        file_name_without_extension = os.path.splitext(file_name)[0]  # Usunięcie rozszerzenia
        tree.insert("", "end", values=(file_name_without_extension,))


def show_image_preview(event):
    """Display image preview when a file is selected from the treeview."""
    selected_item = tree.selection()
    if not selected_item:
        return

    # Pobranie nazwy pliku z wybranego wiersza
    file_name = tree.item(selected_item[0])['values'][0]

    # Ścieżka do pliku obrazu
    image_path = os.path.join("converted_jpgs", file_name + ".jpg")

    try:
        # Otworzenie obrazu
        image = Image.open(image_path)
        image.thumbnail((400, 400))  # Zmniejszenie obrazu do rozmiaru okna podglądu
        image_tk = ImageTk.PhotoImage(image)

        # Wyświetlenie obrazu w oknie
        image_label.config(image=image_tk)
        image_label.image = image_tk
    except Exception as e:
        print(f"Error displaying image: {e}")


def clear_image_preview():
    """Clear the image preview displayed in the window."""
    image_label.config(image=None)
    image_label.image = None


# Etykieta do wyświetlania obrazu
image_label = Label(gui_main_window)
image_label.pack(side="right", padx=20, pady=20)

# Powiązanie kliknięcia w tabeli z funkcją wyświetlania obrazu
tree.bind("<ButtonRelease-1>", show_image_preview)

# Dodanie opcji czyszczenia obrazu po naciśnięciu klawisza
gui_main_window.bind("<Escape>", lambda event: clear_image_preview())

label_view = Label(gui_main_window)
label_view.pack()

gui_main_window.mainloop()
