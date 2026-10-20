# Multi-Format RAG Chatbot

A local Retrieval-Augmented Generation (RAG) chatbot that lets you upload supported documents, extract their text, search relevant passages, and ask questions with source references.

## Features

- **Local answer generation:** Uses Ollama with the configured model (default: `qwen2.5:3b`).
- **Semantic retrieval:** Uses Sentence Transformers (`all-MiniLM-L6-v2`) to create embeddings and ChromaDB to store/search document chunks.
- **Multi-format extraction:** PDF, DOCX, PPTX, XLSX, HTML/HTM, TXT, Markdown, CSV, JSON, XML, common source-code/configuration files, and common image formats.
- **OCR:** Uses Tesseract for images and for PDF pages where normal text extraction returns no readable text.
- **Source references:** Shows the source filename and extracted location for retrieved passages.
- **Streamlit interface:** Runs as a local web application.

## Architecture

1. Upload one or more files in Streamlit.
2. Extract text using format-specific Python libraries; use Tesseract OCR for supported images and scanned PDF pages.
3. Split extracted text into chunks.
4. Generate embeddings with `sentence-transformers/all-MiniLM-L6-v2`.
5. Store chunks, embeddings, and source metadata in a persistent ChromaDB collection (`./chroma_db`).
6. Retrieve relevant passages for a question and send the evidence to Ollama for answer generation.
7. Display the answer and retrieved source passages.

## Requirements

- Windows, macOS, or Linux with a compatible Python installation (the project was developed and tested on Windows).
- Python and the packages listed in `requirements.txt`.
- [Ollama](https://ollama.com/) installed and running locally.
- The configured Ollama model downloaded locally. The default is `qwen2.5:3b`.
- [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) installed for image and scanned-PDF OCR. On the original Windows setup, the executable was at `C:\Program Files\Tesseract-OCR\tesseract.exe`.

## Setup on Windows

Open **Command Prompt (CMD)** and navigate to the project folder:

```cmd
cd /d C:\Users\MOHAMMED\pdf-rag-chatbot
```

Create and activate a virtual environment:

```cmd
python -m venv .venv
.venv\Scripts\activate
```

Install the Python dependencies:

```cmd
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Make sure Ollama is installed and running, then download the default model if needed:

```cmd
ollama pull qwen2.5:3b
```

Run the application:

```cmd
python -m streamlit run app_retrieval_fix.py
```

Streamlit will print a local URL, usually `http://localhost:8501`. Open that URL in your browser.

### Configuration

The app reads optional settings from a `.env` file. Do not commit `.env` or secrets. The defaults in the application are:

```env
OLLAMA_BASE_URL=http://localhost:11434
OLLAMA_MODEL=qwen2.5:3b
OLLAMA_KEEP_ALIVE=30m
TESSERACT_CMD=C:\Program Files\Tesseract-OCR\tesseract.exe
```

Adjust `TESSERACT_CMD` if Tesseract is installed elsewhere. The app uses the default Tesseract path if the variable is not set.

## Supported formats

The application currently lists support for:

- **Documents:** PDF, DOCX, PPTX, XLSX
- **Web and text:** HTML, HTM, TXT, MD
- **Structured data:** CSV, JSON, XML
- **Images:** PNG, JPG/JPEG, BMP, TIF/TIFF, WEBP
- **Code/config/log files:** Python, JavaScript/TypeScript, JSX/TSX, Java, C/C++, C#, Go, Rust, SQL, CSS, YAML, TOML, INI, LOG, shell scripts, BAT, and PowerShell scripts

Actual extraction quality depends on the file contents and format. The app extracts text; it does not perform general visual understanding of charts or diagrams. Embedded images inside Office documents are not currently OCRed. For PDFs, OCR is triggered when a page has no extracted text, so a page that contains both selectable text and image-only text may not have all image text captured.

## Important behavior and limitations

- **Processing a batch replaces the current index.** Upload all files you want to query together before selecting **Process documents**. The current implementation deletes and recreates the `documents` collection when a batch is processed.
- **Local persistence:** ChromaDB data is stored in `./chroma_db` on the local machine. This directory is intentionally excluded from Git.
- **Local model resources:** The embedding model and Ollama model may need to download on first use.
- **OCR quality:** Accuracy depends on image resolution, contrast, font, and language data. English (`eng`) is configured in the current OCR function.
- **Evidence matters:** The chatbot answers from retrieved passages. Review the displayed source passages and citations, especially when multiple files contain similar facts.
- **Privacy:** Documents are processed by this local application and sent to the configured local Ollama service for answer generation. Do not expose the Streamlit server to untrusted networks without adding appropriate security controls.

## Testing notes

The development workflow included controlled tests for PDF extraction, Word, PowerPoint, Excel, CSV, JSON, XML, TXT, Markdown, HTML, source code, image OCR, and scanned-PDF OCR. Known-answer tests passed for the tested examples. These are manual sample tests, not a comprehensive automated test suite or a guarantee for every file in every supported format.

## Repository files

- `app_retrieval_fix.py` — current application entry point.
- `requirements.txt` — Python dependencies.
- `.gitignore` — excludes local environment, secrets, ChromaDB data, backups, and test artifacts.

## Future improvements

- Add automated tests for extraction, retrieval, citation accuracy, and cross-document isolation.
- Pin dependency versions for more reproducible installs.
- Add file-size limits, clearer per-file error reporting, and duplicate-file handling.
- Improve handling of mixed text/image PDF pages and embedded Office images.
- Add an explicit option to append to an existing index instead of always replacing it.

## License

No license has been added yet. If you want others to reuse, modify, or redistribute this project, choose and add an appropriate open-source license.
