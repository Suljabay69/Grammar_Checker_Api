# Grammar Checker API

This project provides an **API for automatic grammar correction** of PDF documents. It converts a PDF to Word, checks and corrects grammar using **OpenAI GPT models**, generates a JSON report of corrections, and outputs a corrected PDF.

---

## Features

- **PDF to Word conversion** (using Microsoft Word automation)
- **Grammar correction** for each paragraph using **OpenAI GPT** (high accuracy, context-aware)
- **JSON report** with detailed token-level diffs and suggestions
- **Word to PDF conversion**
- **REST API** endpoint for easy integration

---

## Requirements

- Python 3.8+
- Microsoft Word (for PDF to Word conversion)
- Windows OS (due to Word automation)
- OpenAI API Key (set as `API_KEY` in your `.env` file)
- The following Python packages (see `requirements.txt`):

  - `fastapi`
  - `uvicorn`
  - `python-docx`
  - `docx2pdf`
  - `pywin32`
  - `openai`
  - `python-dotenv`
  - `asyncio`

**Install all dependencies using:**
```sh
pip install -r requirements.txt
```

**Set your OpenAI API key:**
Create a `.env` file in your project directory with:
```
API_KEY=sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

---

## Usage

### 1. **Start the API server**

```sh
python grammar_checker.py
```

The server will run at `http://localhost:5000`.

---

### 2. **API Endpoint**

**POST** `/api/grammar-check?mode=1&file_code=your_pdf_filename&paragraph_id=1&paragraph_id=2`

- `mode`: `"0"` for full document processing, `"1"` to update only selected paragraphs.
- `file_code`: The name of your PDF file (without `.pdf` extension) located in the `original_pdfs/` folder.
- `paragraph_id`: (Optional, for mode 1) List of paragraph IDs to update.

#### Example Request

```sh
curl -X POST "http://localhost:5000/api/grammar-check?mode=1&file_code=wrong_story&paragraph_id=[1,2]"
```

#### Example Python Request

```python
import requests

params = {
    "mode": "1",
    "file_code": "wrong_story",
    "paragraph_id": ["1", "2"]
}

response = requests.post("http://localhost:5000/api/grammar-check", params=params)
print(response.json())
```

---

### 3. **Response**

```json
{
  "json_filename": "xxx_wrong_story.json",
  "final_pdf_filename": "xxx_wrong_story.pdf",
  "total_improvements": 12,
  "elapsed_time_seconds": 8.34
}
```

- `json_filename`: The JSON report file with grammar corrections details (in `jsons/`).
- `final_pdf_filename`: The corrected PDF file (in `processed_pdfs/`).
- `total_improvements`: Total paragraphs improved and corrected.
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
- The API only accepts **POST** requests.
- Grammar correction is powered by **OpenAI GPT** for high-quality, context-aware proofreading.
- A slight difference between the original PDF and the newly generated PDF can be noticed due to the PDF to Word conversion behaviour.

---

## License

MIT License

---

## Author

Suljabay69