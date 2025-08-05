# PRORDR Maker

**PRORDR Maker** is a desktop application for generating `.txt` files in a specific format, intended for use with a DOS-based system during the creation of electrical projects and implementation of power networks in rural areas. The application simplifies line editing, identifies input errors, and calculates the construction cost in USD.

Developed for **Engeselt Engenharia e Serviços Elétricos**.

---

## ✨ Features

- Easy-to-use GUI built with PySimpleGUI
- Load and save data in `.csv` format
- Export to `.txt` with strict formatting for legacy systems
- Generate project summary as a `.pdf`
- Validate input fields (coordinates, angles, codes)
- Live calculation of construction cost (in USD)

---

## 📁 Output Formats

### `.csv`
Used for temporary saves and continued editing.

### `.txt`
Strictly formatted output required by the DOS system. Includes fixed-width fields and formatted numeric values.

### `.pdf`
Printable project summary with client and surveyor information and tabular data.

---

## 🛠 Technologies Used

- [Python 3.x](https://www.python.org/)
- [PySimpleGUI](https://pypi.org/project/PySimpleGUI/) – GUI framework
- [Pandas](https://pandas.pydata.org/) – for data handling
- [FPDF](https://py-pdf.github.io/fpdf2/) – for PDF exports
- [Tkinter](https://docs.python.org/3/library/tkinter.html) – previously used for cell editing (now deprecated)

---

## 💻 Installation

### Requirements

Install dependencies with:

```bash
pip install PySimpleGUI pandas fpdf openpyxl
