# 🧰 DOCX Metadata & Style Repair Tool  
> **Automated Word Document Fixer** built with Python, Pandoc, and GPT-assisted design.

---

## 📖 Overview
This project demonstrates how to **repair corrupted or misformatted Microsoft Word (`.docx`) files** using Python automation.

It fixes metadata (author, title, subject), cleans XML-level corruption, and rebuilds the file structure by converting it to Markdown and back — producing a **clean, readable, and properly styled DOCX**.

Developed as a demonstration of my ability to handle **document automation and repair workflows** using Python scripting and open-source tools.

---

## ✨ Features
✅ Repair corrupted `.docx` metadata  
✅ Restore missing or blank author/title fields  
✅ Clean broken XML formatting via Markdown conversion  
✅ Automate end-to-end process (input → cleaned output)  
✅ Works on Windows, Linux, and macOS  

---

## 🧠 How It Works
1. Load a DOCX file (even if slightly corrupted).  
2. Extract and rebuild metadata (author, title, etc.) using `python-docx`.  
3. Convert the document to Markdown (`.md`) via **Pandoc**.  
4. Re-export the Markdown back to a clean `.docx` file.  
5. Save the result with repaired metadata and consistent formatting.

---

## 🧩 Project Structure
```
docx_repair_demo/
├── repair_docx.py # main repair script
├── corrupted.docx # sample input (test file)
├── fixed.docx # repaired output
├── clean.md # intermediate markdown conversion
├── requirements.txt # dependencies
└── README.md # this file
```


---

## ⚙️ Tech Stack
- **Python 3.10+**  
- **python-docx** – read and modify Word metadata  
- **Pandoc / pypandoc** – DOCX ↔ Markdown conversion  
- **os / subprocess** – lightweight automation  

---

## 🚀 Usage

🧩 1. Install dependencies
```bash
pip install -r requirements.txt
```
---
🧩 2. Run the repair script
```bash
python repair_docx.py corrupted.docx
```
