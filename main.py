import os
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
    if value is None:
        return ""

    if hasattr(value, "strftime"):
        return value.strftime("%d/%m/%Y")

    return str(value)


def replace_in_paragraph(paragraph, replacements):
    full_text = "".join(run.text for run in paragraph.runs)

    new_text = full_text

    for placeholder, value in replacements.items():
        new_text = new_text.replace(placeholder, value)

    if new_text != full_text:
        if paragraph.runs:
            paragraph.runs[0].text = new_text
            for run in paragraph.runs[1:]:
                run.text = ""


def replace_everywhere(doc, replacements):
    # Κυρίως κείμενο
    for paragraph in doc.paragraphs:
        replace_in_paragraph(paragraph, replacements)

    # Πίνακες
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_in_paragraph(paragraph, replacements)

    # Headers / Footers
    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            replace_in_paragraph(paragraph, replacements)

        for paragraph in section.footer.paragraphs:
            replace_in_paragraph(paragraph, replacements)


def safe_filename(text):
    return re.sub(r'[<>:"/\\|?*]', "_", text)


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

    try:
        wb = load_workbook(excel_path, data_only=True)
        ws = wb.active

        headers = [
            clean_value(cell.value).strip()
            for cell in ws[1]
        ]

        created = 0

        for row_number, row in enumerate(
            ws.iter_rows(min_row=2, values_only=True),
            start=2
        ):
            if all(value is None for value in row):
                continue

            data = {
                headers[i]: clean_value(row[i])
                for i in range(min(len(headers), len(row)))
            }

            replacements = {
                f"[{header}]": value
                for header, value in data.items()
                if header
            }
            # Προσθήκη κενών στα στοιχεία του εκπαιδευτικού
            for field in ["ΕΠΩΝΥΜΟ", "ΟΝΟΜΑ", "ΠΑΤΡΩΝΥΜΟ", "ΚΛΑΔΟΣ"]:
                placeholder = f"[{field}]"

                if placeholder in replacements:
                   replacements[placeholder] = f" {replacements[placeholder]} "

            doc = Document(template_path)
            replace_everywhere(doc, replacements)

            surname = data.get("ΕΠΩΝΥΜΟ", "").strip()
            name = data.get("ΟΝΟΜΑ", "").strip()
            am = data.get("ΑΜ", "").strip()

            if surname or name:
                filename = f"{surname}_{name}"

                if am:
                    filename += f"_{am}"
            else:
                filename = f"ΕΓΓΡΑΦΟ_{row_number}"

            filename = safe_filename(filename) + ".docx"

            output_path = os.path.join(
                output_folder,
                filename
            )

            doc.save(output_path)
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


root = tk.Tk()
root.title("Excel → Word Merge")
root.geometry("650x300")
root.resizable(False, False)

excel_var = tk.StringVar()
template_var = tk.StringVar()
output_var = tk.StringVar()

tk.Label(
    root,
    text="Δημιουργία ατομικών εγγράφων",
    font=("Arial", 16, "bold")
).pack(pady=15)


frame = tk.Frame(root)
frame.pack(fill="x", padx=20)


tk.Button(
    frame,
    text="Επιλογή Excel",
    width=20,
    command=choose_excel
).grid(row=0, column=0, pady=8)

tk.Entry(
    frame,
    textvariable=excel_var,
    width=55
).grid(row=0, column=1, padx=10)


tk.Button(
    frame,
    text="Επιλογή Word Template",
    width=20,
    command=choose_template
).grid(row=1, column=0, pady=8)

tk.Entry(
    frame,
    textvariable=template_var,
    width=55
).grid(row=1, column=1, padx=10)


tk.Button(
    frame,
    text="Φάκελος αποθήκευσης",
    width=20,
    command=choose_output
).grid(row=2, column=0, pady=8)

tk.Entry(
    frame,
    textvariable=output_var,
    width=55
).grid(row=2, column=1, padx=10)


tk.Button(
    root,
    text="ΔΗΜΙΟΥΡΓΙΑ ΕΓΓΡΑΦΩΝ",
    font=("Arial", 12, "bold"),
    width=28,
    height=2,
    command=create_documents
).pack(pady=20)


root.mainloop()
