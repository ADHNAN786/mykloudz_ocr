# 📄 Document OCR to CSV / Excel Converter (Claude Vision & Streamlit)

A standalone Python web application that extracts structured tabular data and line items from PDF documents and images (PNG, JPG, JPEG) using Anthropic's Claude API, allowing direct export to formatted Excel (`.xlsx`) or CSV files.

---

## 🚀 Features

- **Multi-Format Support**: Upload multiple PDFs and document images simultaneously (`.pdf`, `.png`, `.jpg`, `.jpeg`).
- **Native PDF & Vision Support**: Leverages Claude Sonnet 5's native document processing capabilities.
- **Automatic Data Normalization**: Intelligently flattens nested document structures (headers + line items) into tabular rows.
- **Multi-File Provenance**: Includes `source_file` column to easily trace records back to original files.
- **Custom Extraction Hints**: Optional prompt guidance to focus on specific fields (e.g. tax, discounts, receipt numbers).
- **Interactive Data Preview**: View and filter data in real-time with `st.dataframe`.
- **Formatted Export**:
  - **Excel (`.xlsx`)**: Styled headers, zebra borders, formatted numbers, and auto-fitted columns via `openpyxl`.
  - **CSV (`.csv`)**: Clean UTF-8 encoded text export.
- **Raw JSON Inspection**: Collapsible JSON inspector for debugging and reference.

---

## 📦 Setup & Installation

### 1. Clone or navigate to the directory
```bash
cd /home/kunhi/.gemini/antigravity/scratch/claude-doc-ocr
```

### 2. (Optional but recommended) Create a virtual environment
```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Configure API Key
Create a `.env` file in the project root:
```bash
cp .env.example .env
```
Edit `.env` and insert your Anthropic API Key:
```env
ANTHROPIC_API_KEY=sk-ant-api03-...
```
*(You can also enter your key directly inside the app's sidebar during runtime.)*

---

## 🏃 Running the Application

Launch the Streamlit app:
```bash
streamlit run app.py
```

Open your browser at `http://localhost:8501`.
