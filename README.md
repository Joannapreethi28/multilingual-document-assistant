cat > README.md <<'EOF'
# Multilingual Document Assistant using RAG and Lightweight SLM

This project is a multilingual document assistant that helps users upload a document once and perform multiple operations such as document translation, question answering, summarization, key point extraction, and report downloading.

## Features

- Upload PDF, DOCX, or TXT documents
- Translate full document into preferred language
- Side-by-side original and translated document view
- Ask questions from uploaded document
- Generate answers in preferred language
- Display translated and original evidence
- Generate 4-point multilingual summary
- Extract key points
- Download TXT and DOCX reports

## Supported Languages

- English
- Tamil
- Hindi
- French
- Spanish

## Tech Stack

- Python
- Streamlit
- FAISS
- Sentence Transformers
- FLAN-T5 Small
- PyPDF
- python-docx
- deep-translator

## How to Run

```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python -m streamlit run app.py
