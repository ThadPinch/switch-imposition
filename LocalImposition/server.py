import cgi
import json
import mimetypes
import os
import secrets
import tempfile
import time
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

from impose_backend import ImpositionError, parse_config, impose_pdf


ROOT = Path(__file__).resolve().parent
INBOX_DIR = ROOT / "inbox"
STATIC_FILES = {
    "/": ROOT / "index.html",
    "/index.html": ROOT / "index.html",
    "/styles.css": ROOT / "styles.css",
    "/app.js": ROOT / "app.js",
}
DOWNLOADS = {}
DOWNLOAD_TTL_SECONDS = 60 * 60
UPLOAD_COPY_CHUNK_SIZE = 8 * 1024 * 1024


class LocalImpositionHandler(BaseHTTPRequestHandler):
    server_version = "LocalImposition/1.0"

    def do_HEAD(self):
        cleanup_downloads()
        parsed = urlparse(self.path)
        path = unquote(parsed.path)

        if path == "/api/files":
            self.send_json(HTTPStatus.OK, {"files": list_inbox_files()}, head_only=True)
            return

        if path.startswith("/downloads/"):
            token = path.split("/", 2)[-1]
            self.serve_download(token, head_only=True)
            return

        static_path = STATIC_FILES.get(path)
        if static_path is None:
            self.send_error(HTTPStatus.NOT_FOUND, "File not found.")
            return

        self.serve_file(static_path, head_only=True)

    def do_GET(self):
        cleanup_downloads()
        parsed = urlparse(self.path)
        path = unquote(parsed.path)

        if path == "/api/files":
            self.send_json(HTTPStatus.OK, {"files": list_inbox_files()})
            return

        if path.startswith("/downloads/"):
            token = path.split("/", 2)[-1]
            self.serve_download(token)
            return

        static_path = STATIC_FILES.get(path)
        if static_path is None:
            self.send_error(HTTPStatus.NOT_FOUND, "File not found.")
            return

        self.serve_file(static_path)

    def do_POST(self):
        cleanup_downloads()
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        if path != "/api/impose":
            self.send_error(HTTPStatus.NOT_FOUND, "Endpoint not found.")
            return

        try:
            response = self.handle_impose_request()
            self.send_json(HTTPStatus.OK, response)
        except ImpositionError as exc:
            self.send_json(HTTPStatus.BAD_REQUEST, {"error": str(exc)})
        except Exception as exc:
            self.send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"error": str(exc)})

    def handle_impose_request(self):
        content_type = self.headers.get("Content-Type") or ""
        if "multipart/form-data" not in content_type:
            raise ImpositionError("The request must use multipart/form-data.")

        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": content_type,
            },
        )

        config = parse_config(
            {
                key: form.getfirst(key, "")
                for key in form.keys()
                if key not in {"sourcePdf", "inboxFile"}
            }
        )

        inbox_file = form.getfirst("inboxFile", "").strip()
        upload_field = form["sourcePdf"] if "sourcePdf" in form else None
        temp_source_path = None

        if inbox_file:
            source_path = resolve_inbox_file(inbox_file)
            original_name = source_path.name
        elif upload_field is not None and getattr(upload_field, "file", None):
            original_name = Path(upload_field.filename or "imposed.pdf").name
            source_fd, temp_source_path = tempfile.mkstemp(prefix="local-imposition-source-", suffix=".pdf")
            os.close(source_fd)
            with open(temp_source_path, "wb") as destination:
                while True:
                    chunk = upload_field.file.read(UPLOAD_COPY_CHUNK_SIZE)
                    if not chunk:
                        break
                    destination.write(chunk)
            source_path = Path(temp_source_path)
        else:
            raise ImpositionError("Choose a PDF from LocalImposition/inbox or upload one manually.")

        try:
            output_fd, output_path = tempfile.mkstemp(prefix="local-imposition-output-", suffix=".pdf")
            os.close(output_fd)
            try:
                result = impose_pdf(str(source_path), output_path, config)
            except Exception:
                safe_unlink(output_path)
                raise

            output_name = build_download_name(original_name, config.mode, result["cols"], result["rows"])

            token = secrets.token_urlsafe(18)
            DOWNLOADS[token] = {
                "path": output_path,
                "name": output_name,
                "createdAt": time.time(),
            }

            response = dict(result)
            response["downloadUrl"] = f"/downloads/{token}"
            response["outputFileName"] = output_name
            response["sourceName"] = original_name
            return response
        finally:
            if temp_source_path:
                safe_unlink(temp_source_path)

    def serve_file(self, path: Path, head_only: bool = False):
        if not path.exists() or not path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND, "File not found.")
            return

        mime_type = mimetypes.guess_type(str(path))[0] or "application/octet-stream"
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", mime_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        if not head_only:
            self.wfile.write(data)

    def serve_download(self, token: str, head_only: bool = False):
        entry = DOWNLOADS.get(token)
        if not entry:
            self.send_error(HTTPStatus.NOT_FOUND, "Download not found or already used.")
            return

        path = entry["path"]
        try:
            file_size = os.path.getsize(path)
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/pdf")
            self.send_header("Content-Length", str(file_size))
            self.send_header("Content-Disposition", f'attachment; filename="{entry["name"]}"')
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            if not head_only:
                with open(path, "rb") as source:
                    while True:
                        chunk = source.read(UPLOAD_COPY_CHUNK_SIZE)
                        if not chunk:
                            break
                        self.wfile.write(chunk)
        finally:
            if not head_only:
                DOWNLOADS.pop(token, None)
                safe_unlink(path)

    def send_json(self, status: HTTPStatus, payload, head_only: bool = False):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        if not head_only:
            self.wfile.write(body)

    def log_message(self, format, *args):
        return


def cleanup_downloads():
    now = time.time()
    expired_tokens = [
        token
        for token, entry in DOWNLOADS.items()
        if now - entry["createdAt"] > DOWNLOAD_TTL_SECONDS
    ]
    for token in expired_tokens:
        entry = DOWNLOADS.pop(token, None)
        if entry:
            safe_unlink(entry["path"])


def safe_unlink(path: str):
    try:
        os.unlink(path)
    except FileNotFoundError:
        pass


def build_download_name(original_name: str, mode: str, cols: int, rows: int) -> str:
    base_name = Path(original_name).stem
    safe_base = "".join(ch if ch.isalnum() or ch in "._-" else "-" for ch in base_name).strip("-") or "imposed"
    mode_slug = "cut-stack" if mode == "cutAndStack" else "repeat"
    return f"{safe_base}-{mode_slug}-{cols}x{rows}.pdf"


def list_inbox_files():
    ensure_inbox_dir()
    entries = []
    for path in sorted(INBOX_DIR.iterdir()):
        if not path.is_file():
            continue
        if path.suffix.lower() != ".pdf":
            continue
        stat = path.stat()
        entries.append(
            {
                "name": path.name,
                "size": stat.st_size,
                "modifiedAt": int(stat.st_mtime),
            }
        )
    return entries


def resolve_inbox_file(file_name: str) -> Path:
    ensure_inbox_dir()
    safe_name = Path(file_name).name
    if safe_name != file_name:
        raise ImpositionError("Invalid inbox file name.")

    path = INBOX_DIR / safe_name
    if not path.exists() or not path.is_file():
        raise ImpositionError(f"{safe_name} is not in LocalImposition/inbox.")
    if path.suffix.lower() != ".pdf":
        raise ImpositionError("Only PDF files are supported in LocalImposition/inbox.")
    return path


def ensure_inbox_dir():
    INBOX_DIR.mkdir(parents=True, exist_ok=True)


def main():
    ensure_inbox_dir()
    port = int(os.environ.get("LOCAL_IMPOSITION_PORT", "4317"))
    server = ThreadingHTTPServer(("127.0.0.1", port), LocalImpositionHandler)
    print(f"LocalImposition running at http://127.0.0.1:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
