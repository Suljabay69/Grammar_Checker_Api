import re
import os
import time
import json
import difflib
import string
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

def gpt_proofread(text):
    system_msg = (
        "You are a grammar corrector that also classifies changes and suggests synonyms. "
        "You will receive a sentence to correct. Return a JSON object with the following keys:\n"
        "- 'corrected': the corrected version of the sentence\n"
        "- 'changes': a list of changes, where each item includes:\n"
        "  - 'type': one of 'replaced', 'inserted', 'removed', or 'corrected'\n"
        "  - 'original_idx': index in original text (if applicable)\n"
        "  - 'proofread_idx': index in corrected text (if applicable)\n"
        "  - 'original_word': word or punctuation in the original text (if applicable)\n"
        "  - 'proofread_word': corrected word or punctuation (if applicable)\n"
        "  - 'suggestion': up to 3 synonyms (only for 'replaced' or 'inserted')\n"
        "Important: Treat punctuation changes (e.g., commas, periods, quotation marks) as valid changes and include them in the 'changes' list. "
        "Do not ignore them. Every inserted or removed punctuation mark must be reflected in the output."
    )

    user_msg = f"Original sentence:\n{text}\n\nPlease return only the JSON."

    response = client.chat.completions.create(
        model="gpt-4o-mini",
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
    doc = Document(docx_path)
    data = {"pdf_id": pdf_id, "paragraphs": []}
    total_improvements = 0

    for idx, paragraph in enumerate(doc.paragraphs):
        original_text = paragraph.text
        norm_text = original_text
        if not norm_text:
            continue

        try:
            gpt_response = gpt_proofread(norm_text)
        except Exception as e:
            print(f"Error processing paragraph {idx + 1}: {e}")
            continue
        
        corrected_text = gpt_response.get("corrected", norm_text)
        diff_list = gpt_response.get("changes", [])

        if corrected_text != norm_text:
            total_improvements += 1

        original_tokens = norm_text.split()
        proofread_tokens = corrected_text.split()

        original_text_entries = []
        revised_text_entries = []
        print(diff_list)
        for change in diff_list:
            orig_idx = change.get("original_idx")
            proof_idx = change.get("proofread_idx")
            change_type = change.get("type")
            orig_word = change.get("original_word")
            proof_word = change.get("proofread_word")
            suggestions = change.get("suggestion", [])

            if change_type != "inserted" and orig_idx is not None:
                original_text_entries.append({
                    "index": orig_idx,
                    "word": orig_word,
                    "type": "error"
                })

            if proof_idx is not None:
                revised_text_entries.append({
                    "index": proof_idx,
                    "word": proof_word,
                    "type": change_type,
                    "suggestions": suggestions if suggestions else [proof_word]
                })

        para_id = idx + 1
        # Add indexed tokens for original and proofread text
        original_tokens = [{"idx": i, "word": w} for i, w in enumerate(norm_text.split())]
        proofread_tokens = [{"idx": i, "word": w} for i, w in enumerate(corrected_text.split())]

        data["paragraphs"].append({
            "paragraph_id": para_id,
            "original": norm_text,
            "proofread": corrected_text,
            "original_token": original_tokens,
            "proofread_token": proofread_tokens,
            "original_text": original_text_entries,
            "revised_text": revised_text_entries
        })

        # Update paragraph while preserving formatting
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

    return total_improvements


app = FastAPI()

@app.post("/api/grammar-check")
async def main(
    file_code: str = Query(..., description="Filename of the PDF without extension", examples=["wrong_story"]),
):
    start_time = time.time()

    pdf_path = os.path.abspath(f"original_pdfs/{file_code}.pdf")
    docx_path = os.path.abspath(f"parsing_words/{file_code}_temp.docx")
    updated_docx_path = os.path.abspath(f"parsing_words/{file_code}_updated.docx")
    final_pdf_path = os.path.abspath(f"processed_pdfs/xxx_{file_code}.pdf")
    json_output_path = os.path.abspath(f"jsons/xxx_{file_code}.json")

    convert_pdf_to_word(pdf_path, docx_path)
    total_improvements = correct_paragraphs(docx_path, updated_docx_path, json_output_path)
    convert_word_to_pdf(updated_docx_path, final_pdf_path)

    elapsed = time.time() - start_time
    print(f"Total processing time: {elapsed:.2f} seconds")
    print(f"Total Error Found: {total_improvements}")

    return {
        "json_filename": os.path.basename(json_output_path),
        "final_pdf_filename": os.path.basename(final_pdf_path),
        "total_improvements": total_improvements,
        "elapsed_time_seconds": round(elapsed, 2)
    }

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=5000)
