# 💎 YS18 Jewellery Catalog Generator

A powerful **Python GUI application** to generate **brand-style jewellery catalogs (PDF)** using **Excel data + product images**.
Designed for jewellery businesses to quickly create **professional catalogs** with full control over layout, data fields, and styling.

---

## 🚀 Key Features

### 📊 Excel Integration

* Works with **any Excel file**
* Supports **multiple sheets** (user can choose sheet)
* Automatically reads **column headers**
* Handles missing data safely (no `NaN` in PDF)
* Uses **Stock Code** to match images

### 🖼 Image Handling

* Matches images using **Stock Code as filename**
* Supports `.jpg`, `.jpeg`, `.png`
* Ignores Excel rows without images
* Optimized for **large image folders (50,000+ images)**

### 📐 Layout Options

Choose catalog layout:

* **Small** → 5 columns × multiple rows
* **Medium** → 2 columns × multiple rows
* **Large** → 1 column × **2 items per page**

Automatically adjusts spacing and pagination.

---

## 🧾 PDF Customization

### ✅ Select What to Show

User can toggle visibility of:

* Purity
* Gross Weight
* Diamond Weight
* Colour Stone Weight
* Price

> Uncheck → field disappears
> Check again → field reappears

### 🎚 Text Size Control

* Slider to control **PDF text size (70% – 150%)**
* Applies to all selected fields
* Works across all layouts

### 🧠 Smart Formatting

* Numeric values formatted to **2 decimals**
* Price shown without decimals
* Empty fields are skipped automatically
* Clean, brand-style text layout

---

## 🏷 Category Filtering

Automatically groups products by **Stock Code prefix**:

| Category | Prefix |
| -------- | ------ |
| Bracelet | BR     |
| Bangle   | BL     |
| Chain    | C      |
| Earring  | E      |
| Ring     | R      |
| Pendant  | P      |
| Necklace | N      |

User can generate **category-wise PDFs**.

---

## 🧩 Header & Branding

* Custom **PDF header text**
* Optional **logo upload**
* Page numbers included
* Clean white background (brand-ready)

---

## 🖥 GUI Features

* Modern Tkinter-based interface
* Excel selection indicator
* Image folder indicator
* Logo selection indicator
* Progress bar during PDF generation
* Preview PDF before final generation
* Success & error notifications

---

## 📂 Output

* PDFs are saved in:

  ```
  YS18_OUTPUT/
  ```

  (created next to image folder)
* Every run creates a **new PDF**
* Preview files never overwrite final PDFs

---

## 🛠 Installation

### 1️⃣ Install Python Dependencies

```bash
pip install pandas pillow reportlab
```

### 2️⃣ Run the App

```bash
python YS18_Catalog_Generator.py
```

---

## 📦 Build EXE (Windows)

```powershell
python -m PyInstaller --onefile --windowed --name "YS18_Jewellery_Catalog_Generator" YS18_Catalog_Generator.py
```

Output:

```
dist/YS18_Jewellery_Catalog_Generator.exe
```

---

## 📁 Expected Excel Columns (Example)

* Stock Code
* Purity
* Gross Wt
* Total Dia Wt
* Total stone Wt
* Price

> Column names are **selectable inside the app**, so exact naming is flexible.

---

## 🧠 Best Practices

* Keep image names **exactly same as Stock Code**
* Avoid duplicate stock codes in Excel
* Use high-resolution square images for best results

---

## 📌 Use Cases

* Jewellery brand catalogs
* Sales presentations
* WhatsApp / PDF sharing
* Retail & wholesale listings

---

## 🧑‍💻 Author

**YS18 Automation Suite**
Built for scalability, speed, and real business use.

---

If you want:

* Multi-language catalogs
* Watermark support
* Pricing tiers
* Cloud version

👉 Just extend this base.

✨ **Happy Cataloging!**
