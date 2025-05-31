import re
import os
import time 

import win32com.client
from docx import Document
from docx2pdf import convert

import json
import difflib

import language_tool_python
# install java for language_tool_python
# use the link below to install java
# https://www.java.com/en/download/
from nltk.corpus import wordnet
# install bellow items using bash
# python -m nltk.downloader wordnet
# python -m nltk.downloader omw-1.4

from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import JSONResponse
import uvicorn


def convert_pdf_to_word(pdf_path, docx_path):
    """
    Convert a PDF file to a Word (.docx) file using Microsoft Word automation.
    """
    print("Converting PDF to Word using Microsoft Word...")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(pdf_path)
    doc.SaveAs(docx_path, FileFormat=16)  # 16 = wdFormatDocumentDefault (.docx)
    doc.Close()
    word.Quit()
    print("Conversion done.")

def normalize_text(text):
    """
    Normalize text by collapsing whitespace and stripping leading/trailing spaces.
    """
    return re.sub(r'\s+', ' ', text).strip()

def convert_word_to_pdf(updated_docx_path, final_pdf_path):
    """
    Convert a Word (.docx) file back to PDF.
    """
    print("Converting back to PDF...")
    convert(updated_docx_path, final_pdf_path)
    print(f"Final PDF saved as: {final_pdf_path}")
    
def get_synonyms(word, max_synonyms=3):
    """
    Get up to max_synonyms synonyms for a given word using WordNet.
    """
    synonyms = set()
    for syn in wordnet.synsets(word):
        for lemma in syn.lemmas():
            clean = lemma.name().replace('_', ' ')
            if clean.lower() != word.lower():
                synonyms.add(clean)
            if len(synonyms) >= max_synonyms - 1:
                break
        if len(synonyms) >= max_synonyms - 1:
            break
    return [word] + list(synonyms)[:max_synonyms - 1]

def get_diff_and_tokens(original, proofread):
    """
    Compare original and proofread sentences, returning token lists and a diff list.
    The diff list contains details about each change (replace, delete, insert).
    """
    orig_tokens = original.split()
    proof_tokens = proofread.split()
    orig_idx = list(range(len(orig_tokens)))
    proof_idx = list(range(len(proof_tokens)))

    sm = difflib.SequenceMatcher(None, orig_tokens, proof_tokens)
    diff = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'replace':
            for k in range(max(i2 - i1, j2 - j1)):
                o_idx = i1 + k if i1 + k < i2 else i2 - 1
                p_idx = j1 + k if j1 + k < j2 else j2 - 1
                proof_word = proof_tokens[p_idx] if p_idx < len(proof_tokens) else ""
                suggestions = get_synonyms(proof_word)
                diff.append({
                    "original_idx": o_idx,
                    "proofread_idx": p_idx,
                    "original_word": orig_tokens[o_idx] if o_idx < len(orig_tokens) else "",
                    "proofread_word": proof_word,
                    "type": "error",
                    "suggestion": suggestions
                })
        elif tag == 'delete':
            for k in range(i1, i2):
                diff.append({
                    "original_idx": k,
                    "proofread_idx": None,
                    "original_word": orig_tokens[k],
                    "proofread_word": "",
                    "type": "error",
                    "suggestion": []
                })
        elif tag == 'insert':
            for k in range(j1, j2):
                proof_word = proof_tokens[k]
                suggestions = get_synonyms(proof_word)
                diff.append({
                    "original_idx": None,
                    "proofread_idx": k,
                    "original_word": "",
                    "proofread_word": proof_word,
                    "type": "suggestion",
                    "suggestion": suggestions
                })
    return {
        "original": {"text": orig_tokens, "idx": orig_idx},
        "proofread": {"text": proof_tokens, "idx": proof_idx},
        "diff": diff
    }

def correct_paragraphs(docx_path, updated_docx_path, json_output_path, pdf_id="example_pdf_001"):
    """
    Correct grammar in each paragraph of a Word document using LanguageTool,
    update the document, and save a JSON report of all corrections.
    Returns the total number of grammar errors found.
    """
    tool = language_tool_python.LanguageTool('en-US')
    doc = Document(docx_path)
    data = {"pdf_id": pdf_id, "paragraphs": []}
    total_errors = 0

    for idx, paragraph in enumerate(doc.paragraphs):
        original_text = paragraph.text
        norm_text = normalize_text(original_text)
        if not norm_text:
            continue

        matches = tool.check(norm_text)
        total_errors += len(matches)  # Count grammar errors
        corrected_text = language_tool_python.utils.correct(norm_text, matches)

        tokens_and_diff = get_diff_and_tokens(norm_text, corrected_text)
        para_id = f"para_{idx+1:03d}"
        data["paragraphs"].append({
            "paragraph_id": para_id,
            "tokens": {
                "original": tokens_and_diff["original"],
                "proofread": tokens_and_diff["proofread"]
            },
            "diff": tokens_and_diff["diff"]
        })

        # Replace paragraph text with corrected version, preserving formatting
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
    filename: str = Query(..., description="Filename of the PDF without extension", examples="wrong_story"),
):
    """
    Main workflow: PDF to Word, grammar correction, Word to PDF, and JSON report.
    Returns JSON with output filenames, total errors, and elapsed time.
    """
    start_time = time.time()

    # === File paths ===
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

    # Return the required JSON response
    return {
        "json_filename": os.path.basename(json_output_path),
        "final_pdf_filename": os.path.basename(final_pdf_path),
        "total_errors": total_errors,
        "elapsed_time_seconds": round(elapsed, 2)
    }
    
if __name__ == "__main__":
   uvicorn.run(app, host="0.0.0.0", port=5000)