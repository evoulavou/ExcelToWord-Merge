import os
import win32com.client
import re
import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import load_workbook
from docx import Document


def choose_excel():
    path = filedialog.askopenfilename(
        title="Επιλογή αρχείου Excel",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if path:
        excel_var.set(path)


def choose_template():
    path = filedialog.askopenfilename(
        title="Επιλογή προτύπου Word",
        filetypes=[("Word documents", "*.docx")]
    )
    if path:
        template_var.set(path)


def choose_output():
    path = filedialog.askdirectory(
        title="Επιλογή φακέλου αποθήκευσης"
    )
    if path:
        output_var.set(path)


def clean_value(value):
    """
    Μετατρέπει τις τιμές του Excel σε κείμενο.
    Οι ημερομηνίες εμφανίζονται ως ΗΗ/ΜΜ/ΕΕΕΕ.
    """
    if value is None:
        return ""

    if hasattr(value, "strftime"):
        return value.strftime("%d/%m/%Y")

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value)


def safe_filename(text):
    """
    Καθαρίζει χαρακτήρες που δεν επιτρέπονται
    σε ονόματα αρχείων Windows.
    """
    return re.sub(r'[<>:"/\\|?*]', "_", text)


def replace_in_paragraph(paragraph, replacements):
    """
    Αντικατάσταση placeholders μέσα σε paragraph.
    """

    for run in paragraph.runs:
        for placeholder, value in replacements.items():
            if placeholder in run.text:
                run.text = run.text.replace(
                    placeholder,
                    str(value)
                )


def replace_everywhere(doc, replacements):
    """
    Αντικατάσταση placeholders σε όλο το Word:
    - κύριο κείμενο
    - πίνακες
    - headers
    - footers
    """

    # Κύριο κείμενο
    for paragraph in doc.paragraphs:
        replace_in_paragraph(
            paragraph,
            replacements
        )

    # Πίνακες
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_in_paragraph(
                        paragraph,
                        replacements
                    )

    # Headers / Footers
    for section in doc.sections:

        for paragraph in section.header.paragraphs:
            replace_in_paragraph(
                paragraph,
                replacements
            )

        for paragraph in section.footer.paragraphs:
            replace_in_paragraph(
                paragraph,
                replacements
            )


def create_documents():

    excel_path = excel_var.get()
    template_path = template_var.get()
    output_folder = output_var.get()

    if not excel_path or not template_path or not output_folder:
        messagebox.showwarning(
            "Λείπουν στοιχεία",
            "Επίλεξε Excel, Word Template και φάκελο αποθήκευσης."
        )
        return
        
    # Δημιουργία υποφακέλων WORD και PDF
    word_folder = os.path.join(output_folder, "WORD")
    pdf_folder = os.path.join(output_folder, "PDF")

    os.makedirs(word_folder, exist_ok=True)
    os.makedirs(pdf_folder, exist_ok=True)
    
    try:

        wb = load_workbook(
            excel_path,
            data_only=True
        )

        ws = wb.active

        # Πρώτη γραμμή = επικεφαλίδες
        headers = [
            clean_value(cell.value).strip()
            for cell in ws[1]
        ]

        if not any(headers):
            raise Exception(
                "Δεν βρέθηκαν επικεφαλίδες στην πρώτη γραμμή του Excel."
            )

        created = 0

        for row_number, row in enumerate(
            ws.iter_rows(
                min_row=2,
                values_only=True
            ),
            start=2
        ):

            # Αγνοούμε τελείως κενές γραμμές
            if all(value is None for value in row):
                continue

            data = {}

            for i in range(
                min(len(headers), len(row))
            ):

                header = headers[i]

                if not header:
                    continue

                data[header] = clean_value(
                    row[i]
                ).strip()

            # Δημιουργούμε αυτόματα τα placeholders
            #
            # π.χ.
            # Excel: ΕΠΩΝΥΜΟ
            # Word: [ΕΠΩΝΥΜΟ]
            replacements = {
                f"[{header}]": value
                for header, value in data.items()
            }

            # Ανοίγουμε νέο αντίγραφο του template
            doc = Document(
                template_path
            )

            # Αντικατάσταση
            replace_everywhere(
                doc,
                replacements
            )

            # -------------------------
            # Όνομα αρχείου
            # -------------------------

            surname = data.get(
                "ΕΠΩΝΥΜΟ",
                ""
            )

            name = data.get(
                "ΟΝΟΜΑ",
                ""
            )

            am = data.get(
                "ΑΜ",
                ""
            )

            if surname or name:

                filename = f"{surname}_{name}"

                if am:
                    filename += f"_{am}"

            else:

                filename = f"ΕΓΓΡΑΦΟ_{row_number}"

            filename = safe_filename(
                filename
            ) + ".docx"

            output_path = os.path.join(
                output_folder,
                filename
            )

            # Αν υπάρχει ήδη, το αντικαθιστούμε
            if os.path.exists(output_path):
                os.remove(output_path)

            doc.save(
                output_path
            )

            created += 1

        messagebox.showinfo(
            "Ολοκληρώθηκε",
            f"Δημιουργήθηκαν {created} έγγραφα."
        )

    except Exception as e:

        messagebox.showerror(
            "Σφάλμα",
            f"Παρουσιάστηκε σφάλμα:\n\n{e}"
        )


# ============================================================
# GUI
# ============================================================

root = tk.Tk()

root.title(
    "Excel → Word Merge"
)

root.geometry(
    "700x320"
)

root.resizable(
    False,
    False
)


excel_var = tk.StringVar()
template_var = tk.StringVar()
output_var = tk.StringVar()


tk.Label(
    root,
    text="Δημιουργία ατομικών εγγράφων",
    font=("Arial", 16, "bold")
).pack(
    pady=18
)


frame = tk.Frame(
    root
)

frame.pack(
    fill="x",
    padx=20
)


# Excel
tk.Button(
    frame,
    text="Επιλογή Excel",
    width=22,
    command=choose_excel
).grid(
    row=0,
    column=0,
    pady=8
)

tk.Entry(
    frame,
    textvariable=excel_var,
    width=60
).grid(
    row=0,
    column=1,
    padx=10
)


# Word
tk.Button(
    frame,
    text="Επιλογή Word Template",
    width=22,
    command=choose_template
).grid(
    row=1,
    column=0,
    pady=8
)

tk.Entry(
    frame,
    textvariable=template_var,
    width=60
).grid(
    row=1,
    column=1,
    padx=10
)


# Output
tk.Button(
    frame,
    text="Φάκελος αποθήκευσης",
    width=22,
    command=choose_output
).grid(
    row=2,
    column=0,
    pady=8
)

tk.Entry(
    frame,
    textvariable=output_var,
    width=60
).grid(
    row=2,
    column=1,
    padx=10
)


# Δημιουργία
tk.Button(
    root,
    text="ΔΗΜΙΟΥΡΓΙΑ ΕΓΓΡΑΦΩΝ",
    font=("Arial", 12, "bold"),
    width=30,
    height=2,
    command=create_documents
).pack(
    pady=25
)


root.mainloop()
