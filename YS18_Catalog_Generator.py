import os, io, webbrowser
import pandas as pd
from datetime import datetime
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from reportlab.platypus import SimpleDocTemplate, Image, Paragraph, Table, TableStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from PIL import Image as PILImage

APP_NAME = "YS18 Jewellery Catalog Generator"

CATEGORY_PREFIX_MAP = {
    "Bracelet": "BR",
    "Bangle": "BL",
    "Chain": "C",
    "Earring": "E",
    "Ring": "R",
    "Pendant": "P",
    "Necklace": "N"
}

# ---------------- helpers ----------------
def fmt_2(val):
    try:
        if val is None or pd.isna(val) or val == "":
            return None
        return f"{float(val):.2f}"
    except:
        return None

def clean(val):
    if val is None or pd.isna(val) or val == "":
        return None
    return str(val).strip()

# ---------------- PDF ----------------
def generate_pdf(
    excel_file,
    image_folder,
    header_text,
    stock_col,
    selected_headers,
    layout,
    category,
    progress,
    preview
):
    out_dir = os.path.join(os.path.dirname(image_folder), "YS18_OUTPUT")
    os.makedirs(out_dir, exist_ok=True)

    pdf_name = "PREVIEW.pdf" if preview else f"YS18_{datetime.now():%Y%m%d_%H%M%S}.pdf"
    pdf_path = os.path.join(out_dir, pdf_name)

    # ---------- READ EXCEL (ROW 4) ----------
    df = pd.read_excel(excel_file, header=3)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

    if stock_col not in df.columns:
        raise ValueError("Selected Stock Code column not found in Excel")

    df[stock_col] = df[stock_col].astype(str).str.upper().str.strip()
    df = df.drop_duplicates(stock_col, keep="first")
    excel_map = df.set_index(stock_col).to_dict("index")

    # ---------- IMAGE MATCH ----------
    images = []
    for stock in excel_map:
        if category != "All":
            if not stock.startswith(CATEGORY_PREFIX_MAP.get(category, "")):
                continue
        for ext in (".jpg", ".jpeg", ".png", ".JPG", ".PNG"):
            p = os.path.join(image_folder, stock + ext)
            if os.path.exists(p):
                images.append(stock + ext)
                break

    if not images:
        raise ValueError("No matching images found")

    # ---------- LAYOUT ----------
    if layout == "Small":
        COLS, IMG_CM, FONT, ROWS = 5, 3.9, 6, 999
    elif layout == "Medium":
        COLS, IMG_CM, FONT, ROWS = 2, 5.4, 9, 999
    else:
        COLS, IMG_CM, FONT, ROWS = 1, 10.5, 11, 2  # force 2 per page

    IMG_SIZE = IMG_CM * cm

    style = ParagraphStyle(
        "txt",
        fontSize=FONT,
        leading=FONT + 2,
        alignment=1,
        textColor=colors.HexColor("#565656")
    )

    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        leftMargin=0.3 * cm,
        rightMargin=0.3 * cm,
        topMargin=1.3 * cm,
        bottomMargin=0.6 * cm
    )

    elements, page, row = [], [], []
    total = len(images)

    # ---------- BUILD ----------
    for i, img_file in enumerate(images, start=1):
        stock = os.path.splitext(img_file)[0]
        data = excel_map.get(stock, {})

        try:
            pil = PILImage.open(os.path.join(image_folder, img_file)).convert("RGB")
            buf = io.BytesIO()
            pil.save(buf, format="JPEG")
            buf.seek(0)
            img = Image(buf, IMG_SIZE, IMG_SIZE)
        except:
            continue

        blocks = [Paragraph(f"<b>{stock}</b>", style)]

        # -------- LINE 1 --------
        l1 = []
        if "Purity" in selected_headers:
            v = clean(data.get("Purity"))
            if v:
                l1.append(v)
        if "Gross Wt" in selected_headers:
            v = fmt_2(data.get("Gross Wt"))
            if v:
                l1.append(f"G Wt - {v}")
        if "Total Dia Wt" in selected_headers:
            v = fmt_2(data.get("Total Dia Wt"))
            if v:
                l1.append(f"Dia Wt - {v}")

        if l1:
            blocks.append(Paragraph(" | ".join(l1), style))

        # -------- LINE 2 --------
        l2 = []
        if "Total stone Wt" in selected_headers:
            v = fmt_2(data.get("Total stone Wt"))
            if v:
                l2.append(f"Clr Wt - {v}")
        if "Price" in selected_headers:
            v = clean(data.get("Price"))
            if v:
                l2.append(f"Price - {v}")

        if l2:
            blocks.append(Paragraph(" | ".join(l2), style))

        cell = Table([[img]] + [[b] for b in blocks])
        cell.setStyle(TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4)
        ]))

        row.append(cell)

        if len(row) == COLS:
            page.append(row)
            row = []

        if len(page) == ROWS:
            elements.append(Table(
                page,
                colWidths=[(A4[0] - 0.6 * cm) / COLS] * COLS
            ))
            page = []

        progress["value"] = (i / total) * 100
        progress.update()

    if row:
        while len(row) < COLS:
            row.append("")
        page.append(row)

    if page:
        elements.append(Table(
            page,
            colWidths=[(A4[0] - 0.6 * cm) / COLS] * COLS
        ))

    doc.build(elements)

    if preview:
        webbrowser.open(pdf_path)

# ---------------- UI ----------------
class App:
    def __init__(self, root):
        root.title(APP_NAME)
        root.geometry("780x760")
        root.configure(bg="white")

        self.excel = ""
        self.images = ""
        self.header_vars = {}
        self.stock_col = StringVar()

        self.layout = StringVar(value="Small")
        self.category = StringVar(value="All")
        self.header_text = StringVar(value="YS18 Jewellery Catalog Collection")

        Label(root, text=APP_NAME,
              font=("Segoe UI", 16, "bold"),
              bg="white").pack(pady=10)

        Button(root, text="Select Excel", command=self.pick_excel).pack()
        Button(root, text="Select Image Folder", command=self.pick_images).pack()

        Entry(root, textvariable=self.header_text, width=45).pack(pady=4)

        ttk.Combobox(root, values=["Small", "Medium", "Large"],
                     textvariable=self.layout, state="readonly").pack()

        ttk.Combobox(root, values=["All"] + list(CATEGORY_PREFIX_MAP.keys()),
                     textvariable=self.category, state="readonly").pack()

        Label(root, text="Excel Column Mapping",
              font=("Segoe UI", 10, "bold")).pack(pady=6)

        self.headers_frame = Frame(root, bg="white")
        self.headers_frame.pack()

        self.progress = ttk.Progressbar(root, length=420)
        self.progress.pack(pady=10)

        Button(root, text="Preview PDF",
               command=lambda: self.run(True)).pack()
        Button(root, text="Generate Final PDF",
               bg="#d4af37",
               command=lambda: self.run(False)).pack(pady=6)

    def pick_excel(self):
        self.excel = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not self.excel:
            return

        for w in self.headers_frame.winfo_children():
            w.destroy()
        self.header_vars.clear()

        df = pd.read_excel(self.excel, header=3)
        df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

        Label(self.headers_frame, text="Select Stock Code Column",
              bg="white", font=("Segoe UI", 9, "bold")).pack(anchor="w")

        ttk.Combobox(
            self.headers_frame,
            values=list(df.columns),
            textvariable=self.stock_col,
            state="readonly",
            width=30
        ).pack(anchor="w", pady=3)

        self.stock_col.set(df.columns[0])

        Label(self.headers_frame, text="Select Headers to Show in PDF",
              bg="white", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(8, 0))

        for col in df.columns:
            if col == self.stock_col.get():
                continue
            var = BooleanVar(value=True)
            self.header_vars[col] = var
            Checkbutton(self.headers_frame, text=col,
                        variable=var, bg="white").pack(anchor="w")

    def pick_images(self):
        self.images = filedialog.askdirectory()

    def run(self, preview):
        if not self.excel or not self.images:
            messagebox.showerror("Error", "Excel & Image folder required")
            return

        selected_headers = [k for k, v in self.header_vars.items() if v.get()]

        generate_pdf(
            self.excel,
            self.images,
            self.header_text.get(),
            self.stock_col.get(),
            selected_headers,
            self.layout.get(),
            self.category.get(),
            self.progress,
            preview
        )

if __name__ == "__main__":
    root = Tk()
    App(root)
    root.mainloop()

