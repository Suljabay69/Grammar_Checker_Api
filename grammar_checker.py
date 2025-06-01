import re
import os
import time
import json
from typing import List
from dotenv import load_dotenv

import win32com.client
from docx import Document
from docx2pdf import convert

from openai import OpenAI
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import JSONResponse
import uvicorn

# Load environment variables from .env file for API keys, etc.
load_dotenv()
api_key = os.getenv("API_KEY")

# Initialize OpenAI client (not used in main flow, but available)
client = OpenAI(api_key=api_key)

def convert_pdf_to_word(pdf_path, docx_path):
    """
    Convert a PDF file to a Word (.docx) file using Microsoft Word automation.
    """
    print("Converting PDF to Word using Microsoft Word...")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(pdf_path)
    doc.SaveAs(docx_path, FileFormat=16)
    doc.Close()
    word.Quit()
    print("Conversion done.")

def tokenize_with_punctuation(text):
    """
    Tokenizer that splits punctuation as separate tokens.
    """
    return re.findall(r"\w+|[^\w\s]", text, re.UNICODE)

def convert_word_to_pdf(updated_docx_path, final_pdf_path):
    """
    Convert a Word (.docx) file back to PDF.
    """
    print("Converting back to PDF...")
    convert(updated_docx_path, final_pdf_path)
    print(f"Final PDF saved as: {final_pdf_path}")

def gpt_proofread(text):
    """
    Use OpenAI GPT to proofread and classify changes in a sentence.
    Returns a JSON object with original/corrected text, tokens, and changes.
    """
    system_msg = (
        "You are a grammar corrector that also classifies changes and suggests synonyms in a formal manner. "
        "You will receive a sentence to correct. Return a JSON object with the following keys:\n"
        "- 'original': the original input sentence\n"
        "- 'corrected': the corrected version of the sentence\n"
        "- 'original_token': list of objects with 'idx' and 'word' from the original sentence\n"
        "- 'proofread_token': list of objects with 'idx' and 'word' from the corrected sentence\n"
        "- 'changes': a list of changes, where each item includes:\n"
        "  - 'type': one of 'replaced', 'inserted', 'removed', or 'corrected'\n"
        "  - 'original_idx': index in original_token (can be null for insertions)\n"
        "  - 'proofread_idx': index in proofread_token (can be null for removals)\n"
        "  - 'original_word': word or punctuation in the original sentence\n"
        "  - 'proofread_word': corrected word or punctuation\n"
        "  - 'suggestion': up to 3 synonyms (only for 'replaced' or 'inserted')\n"
        "Important: Treat punctuation changes (e.g., commas, periods) as valid changes and reflect them in the tokens and change list."
    )

    user_msg = f"Original sentence:\n{text}\n\nPlease return only the JSON."

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        # model="gpt-4.1",
        messages=[
            {"role": "system", "content": system_msg},
            {"role": "user", "content": user_msg}
        ],
        temperature=0.3,
    )

    content = response.choices[0].message.content.strip()

    try:
        match = re.search(r'\{.*\}', content, re.DOTALL)
        if not match:
            raise ValueError("No JSON object found in GPT response")

        json_str = match.group(0)
        result = json.loads(json_str)

        if "corrected" not in result or "changes" not in result:
            raise ValueError("Missing expected keys in GPT response")

        return result
    except Exception as e:
        print(f"Failed to parse GPT response for input: {text}\nError: {e}")
        raise ValueError(f"Invalid GPT JSON response: {content}")

def correct_paragraphs(docx_path, updated_docx_path, json_output_path, pdf_id="example_pdf_001"):
    """
    Correct grammar in each paragraph of a Word document using GPT,
    update the document, and save a JSON report of all corrections.
    Returns the total number of improved paragraphs.
    """
    doc = Document(docx_path)
    data = {"pdf_id": pdf_id, "paragraphs": []}
    total_improvements = 0

    for idx, paragraph in enumerate(doc.paragraphs):
        original_text = paragraph.text
        if not original_text.strip():
            continue

        try:
            gpt_response = gpt_proofread(original_text)
        except Exception as e:
            print(f"Error processing paragraph {idx + 1}: {e}")
            continue

        corrected_text = gpt_response.get("corrected", original_text)
        if corrected_text != original_text:
            total_improvements += 1

        para_id = idx + 1
        data["paragraphs"].append({
            "paragraph_id": para_id,
            "original": gpt_response.get("original"),
            "proofread": gpt_response.get("corrected"),
            "original_token": gpt_response.get("original_token", []),
            "proofread_token": gpt_response.get("proofread_token", []),
            "original_text": [
                {
                    "index": ch.get("original_idx"),
                    "word": ch.get("original_word"),
                    "type": "error"
                }
                for ch in gpt_response.get("changes", [])
                if ch.get("type") != "inserted" and ch.get("original_idx") is not None
            ],
            "revised_text": [
                {
                    "index": ch.get("proofread_idx"),
                    "word": ch.get("proofread_word"),
                    "type": ch.get("type"),
                    "suggestions": ch.get("suggestion", [ch.get("proofread_word")])
                }
                for ch in gpt_response.get("changes", [])
                if ch.get("proofread_idx") is not None
            ]
        })

        # Update Word paragraph while preserving formatting
        if paragraph.runs:
            ref_run = paragraph.runs[0]
            paragraph.clear()
            new_run = paragraph.add_run(corrected_text)
            new_run.font.name = ref_run.font.name
            new_run.bold = ref_run.bold
            new_run.italic = ref_run.italic
            new_run.underline = ref_run.underline
            new_run.font.size = ref_run.font.size
            if ref_run.font.color and ref_run.font.color.rgb:
                new_run.font.color.rgb = ref_run.font.color.rgb
        else:
            paragraph.text = corrected_text

    doc.save(updated_docx_path)
    print("Document updated with grammar corrections.")

    # Save JSON report
    with open(json_output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"Proofread JSON saved to {json_output_path}")

    return total_improvements

def update_changes_on_pdf(final_pdf_path, updated_docx_path, json_output_path, paragraph_id):
    """
    Updates the paragraphs in a Word document based on proofread data from a JSON file.
    Only updates paragraphs whose IDs are in paragraph_id.
    Returns the number of paragraphs updated.
    """
    # Load proofreading results from JSON
    with open(json_output_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    
    paragraphs_data = data.get("paragraphs", [])
    paragraph_id_set = set(paragraph_id)
    
    # Load the existing DOCX file
    doc = Document(updated_docx_path)
    updated_count = 0

    for para in paragraphs_data:
        pid = str(para.get("paragraph_id"))
        if pid in paragraph_id_set:
            print(pid)
            proofread_text = para.get("proofread")
            para_index = int(pid) - 1  # Adjust for 0-based indexing

            if 0 <= para_index < len(doc.paragraphs):
                paragraph = doc.paragraphs[para_index]
                if paragraph.runs:
                    ref_run = paragraph.runs[0]
                    paragraph.clear()
                    new_run = paragraph.add_run(proofread_text)
                    new_run.font.name = ref_run.font.name
                    new_run.bold = ref_run.bold
                    new_run.italic = ref_run.italic
                    new_run.underline = ref_run.underline
                    new_run.font.size = ref_run.font.size
                    if ref_run.font.color and ref_run.font.color.rgb:
                        new_run.font.color.rgb = ref_run.font.color.rgb
                else:
                    paragraph.text = proofread_text
                updated_count += 1

    # Save changes to DOCX
    doc.save(updated_docx_path)

    return updated_count

# Initialize FastAPI app
app = FastAPI()

@app.post("/api/grammar-check")
async def main(
    mode: str = Query(..., description="Mode of Processing", examples={"mode": ["1"]}),
    file_code: str = Query(..., description="Filename of the PDF without extension", examples={"file_code": ["wrong_story"]}),
    paragraph_id: List[str] = Query(..., description="IDs of the changed paragraphs", examples={"paragraph_id": ["1", "2"]})
):
    """
    Main API endpoint for grammar checking and PDF processing.
    - mode="0": Full process (PDF→Word→Proofread→PDF)
    - mode="1": Update only selected paragraphs (using paragraph_id)
    Returns output filenames, total improvements, and elapsed time.
    """
    start_time = time.time()

    # Build file paths based on file_code
    pdf_path = os.path.abspath(f"original_pdfs/{file_code}.pdf")
    docx_path = os.path.abspath(f"parsing_words/{file_code}_temp.docx")
    updated_docx_path = os.path.abspath(f"parsing_words/{file_code}_updated.docx")
    final_pdf_path = os.path.abspath(f"processed_pdfs/xxx_{file_code}.pdf")
    json_output_path = os.path.abspath(f"jsons/xxx_{file_code}.json")

    # Main processing logic based on mode
    if mode == "0":
        # Full process: convert, proofread, and save all
        convert_pdf_to_word(pdf_path, docx_path)
        total_improvements = correct_paragraphs(docx_path, updated_docx_path, json_output_path)
        convert_word_to_pdf(updated_docx_path, final_pdf_path)
    else:
        # Only update selected paragraphs
        convert_pdf_to_word(final_pdf_path, docx_path)
        total_improvements = update_changes_on_pdf(final_pdf_path, updated_docx_path, json_output_path, paragraph_id)
        convert_word_to_pdf(updated_docx_path, final_pdf_path)

    elapsed = time.time() - start_time
    print(f"Total processing time: {elapsed:.2f} seconds")
    print(f"Total Error Found: {total_improvements}")

    # Return summary as JSON
    return {
        "json_filename": os.path.basename(json_output_path),
        "final_pdf_filename": os.path.basename(final_pdf_path),
        "total_improvements": total_improvements,
        "elapsed_time_seconds": round(elapsed, 2)
    }

if __name__ == "__main__":
    # Run the FastAPI app with Uvicorn
    uvicorn.run(app, host="0.0.0.0", port=5000)