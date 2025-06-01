import re
import os
import time
import json
import difflib
from dotenv import load_dotenv

import win32com.client
from docx import Document
from docx2pdf import convert

from openai import OpenAI
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import JSONResponse
import uvicorn

# Load environment variables from .env
load_dotenv()
api_key = os.getenv("API_KEY")

client = OpenAI(api_key=api_key)

def convert_pdf_to_word(pdf_path, docx_path):
    print("Converting PDF to Word using Microsoft Word...")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(pdf_path)
    doc.SaveAs(docx_path, FileFormat=16)
    doc.Close()
    word.Quit()
    print("Conversion done.")

def normalize_text(text):
    return re.sub(r'\s+', ' ', text).strip()

def convert_word_to_pdf(updated_docx_path, final_pdf_path):
    print("Converting back to PDF...")
    convert(updated_docx_path, final_pdf_path)
    print(f"Final PDF saved as: {final_pdf_path}")

def get_diff(original, corrected):
    orig_tokens = original.split()
    corr_tokens = corrected.split()
    sm = difflib.SequenceMatcher(None, orig_tokens, corr_tokens)
    diff = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag in ('replace', 'delete', 'insert'):
            diff.append({"tag": tag, "original": orig_tokens[i1:i2], "corrected": corr_tokens[j1:j2]})
    return diff

def gpt_proofread(text):
    prompt = (
        "Correct the grammar of the following sentence. After that, provide up to 3 synonyms for each changed word.\n\n"
        f"Original: {text}\n\nRespond in JSON like this:\n"
        '{{"corrected": "...", "synonyms": {{"changed_word": ["syn1", "syn2", "syn3"]}}}}'
    )

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are a grammar corrector and synonym suggester."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.3,
    )

    content = response.choices[0].message.content.strip()

    print("GPT raw response:", content)

    # Try to extract the first JSON object using regex
    try:
        match = re.search(r'\{.*\}', content, re.DOTALL)
        if not match:
            raise ValueError("No JSON object found in GPT response")

        json_str = match.group(0)
        result = json.loads(json_str)

        if "corrected" not in result or "synonyms" not in result:
            raise ValueError("Missing expected keys in GPT response")

        return result
    except Exception as e:
        print(f"Failed to parse GPT response for input: {text}\nError: {e}")
        raise ValueError(f"Invalid GPT JSON response: {content}")

def correct_paragraphs(docx_path, updated_docx_path, json_output_path, pdf_id="example_pdf_001"):
    doc = Document(docx_path)
    data = {"pdf_id": pdf_id, "paragraphs": []}
    total_errors = 0

    for idx, paragraph in enumerate(doc.paragraphs):
        original_text = paragraph.text
        norm_text = normalize_text(original_text)
        if not norm_text:
            continue

        try:
            gpt_response = gpt_proofread(norm_text)
        except Exception as e:
            print(f"Error processing paragraph {idx + 1}: {e}")
            continue

        corrected_text = gpt_response.get("corrected", norm_text)
        synonyms = gpt_response.get("synonyms", {})

        if corrected_text != norm_text:
            total_errors += 1

        diff = get_diff(norm_text, corrected_text)
        para_id = f"para_{idx+1:03d}"
        data["paragraphs"].append({
            "paragraph_id": para_id,
            "original": norm_text,
            "corrected": corrected_text,
            "diff": diff,
            "synonyms": synonyms
        })

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

    with open(json_output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"Proofread JSON saved to {json_output_path}")

    return total_errors

app = FastAPI()

@app.post("/api/grammar-check")
async def main(
    filename: str = Query(..., description="Filename of the PDF without extension", examples=["wrong_story"]),
):
    start_time = time.time()

    pdf_path = os.path.abspath(f"original_pdfs/{filename}.pdf")
    docx_path = os.path.abspath(f"parsing_words/{filename}_temp.docx")
    updated_docx_path = os.path.abspath(f"parsing_words/{filename}_updated.docx")
    final_pdf_path = os.path.abspath(f"processed_pdfs/xxx_{filename}.pdf")
    json_output_path = os.path.abspath(f"jsons/xxx_{filename}.json")

    convert_pdf_to_word(pdf_path, docx_path)
    total_errors = correct_paragraphs(docx_path, updated_docx_path, json_output_path)
    convert_word_to_pdf(updated_docx_path, final_pdf_path)

    elapsed = time.time() - start_time
    print(f"Total processing time: {elapsed:.2f} seconds")
    print(f"Total Error Found: {total_errors}")

    return {
        "json_filename": os.path.basename(json_output_path),
        "final_pdf_filename": os.path.basename(final_pdf_path),
        "total_errors": total_errors,
        "elapsed_time_seconds": round(elapsed, 2)
    }

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=5000)
