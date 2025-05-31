# Grammar Checker API

This project provides an **API for automatic grammar correction** of PDF documents. It converts a PDF to Word, checks and corrects grammar using [LanguageTool](https://languagetool.org/), generates a JSON report of corrections, and outputs a corrected PDF.

---

## Features

- **PDF to Word conversion** (using Microsoft Word automation)
- **Grammar correction** for each paragraph (using `language-tool-python`)
- **JSON report** with detailed token-level diffs and suggestions
- **Word to PDF conversion**
- **REST API** endpoint for easy integration

---

## Requirements

- Python 3.8+
- Microsoft Word (for PDF to Word conversion)
- Java (required by `language-tool-python`)
- [LanguageTool](https://languagetool.org/) (automatically handled by `language-tool-python`)
- Windows OS (due to Word automation)
- The following Python packages (see `requirements.txt`):

  - `fastapi`
  - `uvicorn`
  - `python-docx`
  - `docx2pdf`
  - `language-tool-python`
  - `nltk`
  - `pywin32`

**Install all dependencies using**:
```sh
pip install -r requirements.txt
python -m nltk.downloader wordnet
python -m nltk.downloader omw-1.4
```

**Install Java:**  
Download and install from [https://www.java.com/en/download/](https://www.java.com/en/download/)

---

## Usage

### 1. **Start the API server**

```sh
python grammar_checker.py
```

The server will run at `http://localhost:5000`.

---

### 2. **API Endpoint**

**POST** `/api/grammar-check?filename=your_pdf_filename`

- `filename`: The name of your PDF file (without `.pdf` extension) located in the `original_pdfs/` folder.

#### Example Request

```sh
curl -X POST "http://localhost:5000/api/grammar-check?filename=wrong_story"
```

#### Example Python Request

```python
import requests
response = requests.post("http://localhost:5000/api/grammar-check", params={"filename": "wrong_story"})
print(response.json())
```

---

### 3. **Response**

```json
{
  "json_filename": "xxx_wrong_story.json",
  "final_pdf_filename": "xxx_wrong_story.pdf",
  "total_errors": 12,
  "elapsed_time_seconds": 8.34
}
```

- `json_filename`: The JSON report file with grammar corrections details (in `jsons/`).
- `final_pdf_filename`: The corrected PDF file (in `processed_pdfs/`).
- `total_errors`: Total grammar errors found and corrected.
- `elapsed_time_seconds`: Total processing time.

---

## Folder Structure

```
original_pdfs/        # Place your input PDFs here
parsing_words/        # Temporary and updated Word files
processed_pdfs/       # Output corrected PDFs
jsons/                # Output JSON reports
grammar_checker.py    # Main API and logic
```

---

## Notes

- This project is designed for Windows due to Microsoft Word automation.
- Make sure Java is installed and available in your PATH for `language-tool-python` to work.
- The API only accepts **POST** requests.
- A slight difference between the original pdf and the newly generated pdf can be noticed due to the PDF to Word conversion behaviour.

---

## License

MIT License

---

## Author

Suljabay69
