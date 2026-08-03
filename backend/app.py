"""
Content Check backend — opt-in enhanced parsing.

This exists purely to replace apps/contentcheck/js/parser.js's in-browser
parsing with python-pptx / python-docx / pdfplumber for documents where the
hand-rolled JS OOXML/pdf.js parsing struggles (SmartArt edge cases, complex
PDF tables). It is NOT required for Content Check to work — the frontend
still runs fully offline by default; a user must explicitly enable the
"enhanced parsing" toggle, which is the only thing that causes a file to
leave the browser.

No uploaded file content is logged or persisted anywhere — files are parsed
in memory and discarded once the response is sent.
"""
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware

from parsers import parse_file

app = FastAPI(title="Content Check Parsing Backend")

# Wide open by design: this is a stateless, read-only parsing endpoint with
# no auth/session/cookie concept, called directly from a static site that
# can be hosted from any origin (GitHub Pages, Render, a local dev server).
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/parse")
async def parse(file: UploadFile = File(...)):
    filename = file.filename or "upload"
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""

    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")

    try:
        result = parse_file(ext, data, filename)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Could not parse {filename}: {e}")

    return result
