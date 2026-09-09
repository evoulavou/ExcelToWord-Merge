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
    """
    Μετατρέπει τις τιμές του Excel σε κείμενο.
    Οι ημερομηνίες εμφανίζονται ως ΗΗ/ΜΜ/ΕΕΕΕ.
    """
    if value is None:
        return ""

    if hasattr(value, "strftime"):
        return value.strftime("%d/%m/%Y")

    # Αποφυγή π.χ. 111111.0 όταν το Excel έχει αριθμό
    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value)


def replace_in_paragraph(paragraph, replacements):
    """
    Αντικαθιστά placeholders όπως [ΕΠΩΝΥΜΟ]
    ακόμη κι αν το Word τα έχει χωρίσει σε πολλά runs.

    Δεν ξαναγράφει ολόκληρη την παράγραφο,
    ώστε να διατηρούνται όσο γίνεται η μορφοποίηση,
    η σειρά του κειμένου και τα κενά.
    """

    for placeholder, value in replacements.items():

        while placeholder in paragraph.text:

            full_text = "".join(run.text for run in paragraph.runs)

            start = full_text.find(placeholder)

            if start == -1:
                break

            end = start + len(placeholder)

            current_pos = 0

            start_run = None
            start_offset = None

            end_run = None
            end_offset = None

            for i, run in enumerate(paragraph.runs):

                run_start = current_pos
                run_end = current_pos + len(run.text)

                if start_run is None and start < run_end:
                    start_run = i
                    start_offset = start - run_start

                if end_run is None and end <= run_end:
                    end_run = i
                    end_offset = end - run_start
                    break

                current_pos = run_end

            if start_run is None or end_run is None:
                break

            # Το placeholder βρίσκεται ολόκληρο σε ένα run
            if start_run == end_run:

                run = paragraph.runs[start_run]

                before = run.text[:start_offset]
                after = run.text[end_offset:]

                run.text = before + str(value) + after

            # Το placeholder είναι σπασμένο σε πολλά runs
            else:

                first_run = paragraph.runs[start_run]
                last_run = paragraph.runs[end_run]

                before = first_run.text[:start_offset]
                after = last_run.text[end_offset:]

                first_run.text = before + str(value)

                # Καθαρίζουμε τα ενδιάμεσα runs
                for i in range(start_run + 1, end_run):
                    paragraph.runs[i].text = ""

                # Κρατάμε οτιδήποτε υπήρχε μετά το placeholder
                last_run.text = after


def replace_everywhere(doc, replacements):
    """
    Αντικατάσταση placeholders σε:
    - κύριο κείμενο
    - πίνακες
    - κεφαλίδες
    - υποσέλιδα
    """

    # Κύριο κείμενο
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
    """
    Αφαιρεί χαρακτήρες που δεν επιτρέπονται
    στα ονόματα αρχείων των Windows.
    """
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

        wb = load_workbook(
            excel_path,
            data_only=True
        )

        ws = wb.active

        # Διαβάζουμε τις επικεφαλίδες της πρώτης γραμμής
        headers = [
            clean_value(cell.value).strip()
            for cell in ws[1]
        ]

        created = 0

        for row_number, row in enumerate(
            ws.iter_rows(
                min_row=2,
                values_only=True
            ),
            start=2
        ):

            # Αγνοούμε εντελώς κενές γραμμές
            if all(value is None for value in row):
                continue

            data = {}

            for i in range(min(len(headers), len(row))):

                header = headers[i]

                if not header:
                    continue

                data[header] = clean_value(row[i])

            # Δημιουργούμε αυτόματα τα placeholders.
            #
            # Π.χ.
            # Excel: ΕΠΩΝΥΜΟ
            # Word: [ΕΠΩΝΥΜΟ]
            replacements = {
                f"[{header}]": value
                for header, value in data.items()
            }

            # Φορτώνουμε νέο αντίγραφο του template
            # για κάθε εκπαιδευτικό
            doc = Document(template_path)

            replace_everywhere(
                doc,
                replacements
            )

            # ------------------------------
            # Όνομα αρχείου
            # ------------------------------

            surname = data.get(
                "ΕΠΩΝΥΜΟ",
                ""
            ).strip()

            name = data.get(
                "ΟΝΟΜΑ",
                ""
            ).strip()

            am = data.get(
                "ΑΜ",
                ""
            ).strip()

            if surname or name:

                filename = f"{surname}_{name}"

                if am:
                    filename += f"_{am}"

            else:

                filename = f"ΕΓΓΡΑΦΟ_{row_number}"

            filename = safe_filename(
                filename
            )

            filename += ".docx"

            output_path = os.path.join(
                output_folder,
                filename
            )

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
