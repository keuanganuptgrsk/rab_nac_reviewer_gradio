from pathlib import Path


def _easyocr_text(image_path):
    import easyocr

    reader = easyocr.Reader(["id", "en"], gpu=False)
    result = reader.readtext(str(image_path), detail=0)
    return "\n".join(result)


def _paddleocr_text(image_path):
    from paddleocr import PaddleOCR

    ocr = _build_paddleocr()
    result = ocr.ocr(str(image_path), cls=True)
    lines = []
    for page in result or []:
        if isinstance(page, dict):
            texts = page.get("rec_texts") or page.get("texts") or []
            lines.extend(str(text) for text in texts if text)
            continue
        for item in page or []:
            if item and len(item) > 1:
                lines.append(str(item[1][0]))
    return "\n".join(lines)


def _build_paddleocr():
    from paddleocr import PaddleOCR

    attempts = [
        {"use_angle_cls": True, "lang": "latin", "show_log": False},
        {"use_angle_cls": True, "lang": "en", "show_log": False},
        {"lang": "latin"},
        {"lang": "en"},
    ]
    last_error = None
    for kwargs in attempts:
        try:
            return PaddleOCR(**kwargs)
        except Exception as exc:
            last_error = exc
    raise last_error


def _tesseract_text(image_path):
    import pytesseract
    from PIL import Image

    return pytesseract.image_to_string(Image.open(image_path), lang="ind+eng")


def extract_text_from_image(image_path, mode="auto"):
    path = Path(image_path)
    errors = []
    engines = ["paddleocr", "easyocr", "tesseract"] if mode in ("auto", "", None) else [mode]
    for engine in engines:
        if engine == "disabled":
            return "", "OCR dinonaktifkan."
        try:
            if engine == "easyocr":
                text = _easyocr_text(path)
            elif engine == "paddleocr":
                text = _paddleocr_text(path)
            elif engine == "tesseract":
                text = _tesseract_text(path)
            else:
                continue
            if text.strip():
                return text, f"OCR berhasil menggunakan {engine}."
        except Exception as exc:
            errors.append(f"{engine}: {exc}")
    return "", "OCR tidak tersedia/berhasil. Upload Excel/CSV atau PDF berbasis teks. " + " | ".join(errors[:3])


def extract_text_from_pdf_scan(pdf_path, mode="auto", max_pages=5):
    if mode == "disabled":
        return "", "OCR dinonaktifkan."
    try:
        import fitz
    except Exception as exc:
        return "", f"PyMuPDF tidak tersedia untuk render OCR: {exc}"
    texts, notes = [], []
    with fitz.open(pdf_path) as doc:
        for i, page in enumerate(doc[:max_pages]):
            pix = page.get_pixmap(dpi=160)
            tmp = Path(pdf_path).with_suffix(f".page_{i+1}.png")
            pix.save(tmp)
            text, note = extract_text_from_image(tmp, mode)
            texts.append(text)
            notes.append(f"Halaman {i+1}: {note}")
            try:
                tmp.unlink()
            except Exception:
                pass
    return "\n".join(texts), " ".join(notes)
