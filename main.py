import os
import pyautogui
import tkinter as tk
from tkinter import filedialog, Label, Menu, ttk
from pdf2image import convert_from_path
from PIL import Image, ImageTk, ImageDraw

# Lista do przechowywania nazw plików, które zostały przekonwertowane
converted_files = []

# Zmienna do przechowywania stanu rysowania
is_drawing = False
start_x, start_y = 0, 0
rect_image = None
draw = None

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
        global rect_image, draw
        rect_image = Image.open(image_path)
        rect_image.thumbnail((400, 400))  # Zmniejszenie obrazu do rozmiaru okna podglądu
        image_tk = ImageTk.PhotoImage(rect_image)

        # Inicjalizacja rysowania
        draw = ImageDraw.Draw(rect_image)

        # Wyświetlenie obrazu w oknie
        image_label.config(image=image_tk)
        image_label.image = image_tk
    except Exception as e:
        print(f"Error displaying image: {e}")

def clear_image_preview():
    """Clear the image preview displayed in the window."""
    image_label.config(image=None)
    image_label.image = None

def start_drawing(event):
    """Start drawing a rectangle."""
    global is_drawing, start_x, start_y
    if not rect_image:
        return

    is_drawing = True
    start_x, start_y = event.x, event.y

def draw_rectangle(event):
    """Draw the rectangle on the image preview."""
    global is_drawing, start_x, start_y, draw
    if is_drawing and rect_image:
        # Rysowanie prostokąta
        temp_image = rect_image.copy()
        temp_draw = ImageDraw.Draw(temp_image)
        temp_draw.rectangle([start_x, start_y, event.x, event.y], outline="red", width=3)

        # Przełącz obraz na tymczasowy
        temp_image_tk = ImageTk.PhotoImage(temp_image)
        image_label.config(image=temp_image_tk)
        image_label.image = temp_image_tk

def stop_drawing(event):
    """Stop drawing the rectangle."""
    global is_drawing, draw
    if is_drawing:
        is_drawing = False
        # Narysuj ostateczny prostokąt
        if rect_image:
            draw.rectangle([start_x, start_y, event.x, event.y], outline="red", width=3)
            image_tk = ImageTk.PhotoImage(rect_image)
            image_label.config(image=image_tk)
            image_label.image = image_tk

def enable_drawing():
    """Enable drawing mode."""
    global is_drawing
    is_drawing = True

# Etykieta do wyświetlania obrazu
image_label = Label(gui_main_window)
image_label.pack(side="right", padx=20, pady=20)

# Powiązanie kliknięcia w tabeli z funkcją wyświetlania obrazu
tree.bind("<ButtonRelease-1>", show_image_preview)

# Przycisk do włączenia trybu rysowania
draw_button = tk.Button(gui_main_window, text="Draw Rectangle", command=enable_drawing)
draw_button.pack(side="bottom", pady=10)

# Rysowanie na obrazku
image_label.bind("<ButtonPress-1>", start_drawing)
image_label.bind("<B1-Motion>", draw_rectangle)
image_label.bind("<ButtonRelease-1>", stop_drawing)

gui_main_window.mainloop()
