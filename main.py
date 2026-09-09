import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import load_workbook
import win32com.client as win32


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

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value)


def safe_filename(text):
    return re.sub(r'[<>:"/\\|?*]', "_", text)


def replace_word_text(doc, placeholder, value):
    """
    Χρησιμοποιεί το ίδιο το Microsoft Word για Find & Replace.
    Έτσι διατηρούνται μορφοποίηση, στοίχιση, κενά κτλ.
    """

    for story_range in doc.StoryRanges:

        current_range = story_range

        while current_range is not None:

            find = current_range.Find

            find.ClearFormatting()
            find.Replacement.ClearFormatting()

            find.Text = placeholder
            find.Replacement.Text = str(value)

            find.Forward = True
            find.Wrap = 1
            find.Format = False
            find.MatchCase = False
            find.MatchWholeWord = False

            find.Execute(
                Replace=2
            )

            try:
                current_range = current_range.NextStoryRange
            except:
                current_range = None


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

    word = None

    try:

        wb = load_workbook(
            excel_path,
            data_only=True
        )

        ws = wb.active

        headers = [
            clean_value(cell.value).strip()
            for cell in ws[1]
        ]

        word = win32.DispatchEx("Word.Application")

        word.Visible = False
        word.DisplayAlerts = False

        created = 0

        for row_number, row in enumerate(
            ws.iter_rows(
                min_row=2,
                values_only=True
            ),
            start=2
        ):

            if all(value is None for value in row):
                continue

            data = {}

            for i in range(min(len(headers), len(row))):

                header = headers[i]

                if not header:
                    continue

                data[header] = clean_value(row[i])

            # Ανοίγουμε κάθε φορά το αρχικό template
            doc = word.Documents.Open(
                os.path.abspath(template_path)
            )

            # Αντικατάσταση όλων των placeholders
            for header, value in data.items():

                placeholder = f"[{header}]"

                replace_word_text(
                    doc,
                    placeholder,
                    value
                )

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
            ) + ".docx"

            output_path = os.path.abspath(
                os.path.join(
                    output_folder,
                    filename
                )
            )

            # 16 = DOCX
            doc.SaveAs2(
                output_path,
                FileFormat=16
            )

            doc.Close(
                SaveChanges=False
            )

            created += 1

        word.Quit()
        word = None

        messagebox.showinfo(
            "Ολοκληρώθηκε",
            f"Δημιουργήθηκαν {created} έγγραφα."
        )

    except Exception as e:

        if word is not None:
            try:
                word.Quit()
            except:
                pass

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


frame = tk.Frame(root)

frame.pack(
    fill="x",
    padx=20
)


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
