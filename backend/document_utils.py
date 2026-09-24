import io
import re


def _strip_html_tags(html_text: str) -> str:
    # <script>と<style>の中身は読み取っても意味がないので丸ごと除去
    html_text = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", html_text, flags=re.DOTALL | re.IGNORECASE)
    # 残りのタグを除去
    text = re.sub(r"<[^>]+>", " ", html_text)
    # 連続する空白をまとめる
    return re.sub(r"[ \t]+", " ", text).strip()


def extract_text_from_file(file_bytes: bytes, filename: str) -> str:
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""

    if ext in ("txt", "md", "json"):
        return file_bytes.decode("utf-8", errors="ignore")

    if ext in ("html", "htm"):
        raw_html = file_bytes.decode("utf-8", errors="ignore")
        return _strip_html_tags(raw_html)

    if ext == "pdf":
        import pdfplumber
        text_parts = []
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text_parts.append(page_text)
        return "\n".join(text_parts)

    raise ValueError(f"Unsupported file type: .{ext}")


MAX_EXTRACTED_CHARS = 60000


def truncate_text(text: str, max_chars: int = MAX_EXTRACTED_CHARS) -> str:
    if len(text) <= max_chars:
        return text
    return text[:max_chars] + "\n\n...(以下省略、文字数上限のため切り捨てられました)"
