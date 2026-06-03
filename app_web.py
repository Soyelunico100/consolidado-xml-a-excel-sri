import html
import mimetypes
import os
import shutil
import subprocess
import sys
import threading
import traceback
import warnings
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse

warnings.filterwarnings("ignore", category=DeprecationWarning, module="cgi")
import cgi

import consolidado_xml_a_excel as converter


HOST = "127.0.0.1"
PORT = 8765
BASE_DIR = Path(__file__).resolve().parent


def ensure_folders():
    for folder in (
        converter.XML_FOLDER,
        converter.SALIDA_EXCEL_DIR,
        converter.PROCESADOS_DIR,
        converter.ERRORES_DIR,
        converter.PDF_ROOT_DIR,
    ):
        folder.mkdir(parents=True, exist_ok=True)


def list_files(folder, pattern):
    if not folder.exists():
        return []
    return sorted(folder.glob(pattern), key=lambda item: item.stat().st_mtime, reverse=True)


def safe_child(base, name):
    base = base.resolve()
    target = (base / name).resolve()
    if base != target and base not in target.parents:
        raise ValueError("Ruta no permitida")
    return target


def unique_path(path):
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    counter = 2
    while True:
        candidate = path.with_name(f"{stem}_{counter}{suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def open_in_explorer(path):
    try:
        if path.is_file():
            subprocess.Popen(["explorer", "/select,", str(path)])
        else:
            subprocess.Popen(["explorer", str(path)])
    except Exception:
        pass


def activar_si_hace_falta(clave):
    if not getattr(converter, "REQUIERE_ACTIVACION", False):
        return True, ""
    script_hash = converter.calcular_hash_script()
    if converter.activacion_valida(script_hash):
        return True, ""
    clave = (clave or "").strip().strip('"').strip("'")
    clave = clave.replace("\ufeff", "").replace("\u200b", "")
    if not clave:
        return False, "Escribe la clave de instalacion para activar esta computadora."
    import hashlib

    clave_hash = hashlib.sha256(clave.encode("utf-8")).hexdigest()
    if clave_hash != converter.CLAVE_ACTIVACION_HASH_SHA256:
        return False, "Clave incorrecta. No se genero ningun archivo."
    converter.guardar_activacion(script_hash)
    return True, "Activacion correcta."


def run_converter(clave=""):
    ok, activation_message = activar_si_hace_falta(clave)
    if not ok:
        return activation_message

    old_stdout = sys.stdout
    old_stderr = sys.stderr
    from io import StringIO

    output = StringIO()
    sys.stdout = output
    sys.stderr = output
    try:
        converter.main()
    except SystemExit as exc:
        if exc.code:
            print(f"Proceso detenido con codigo: {exc.code}")
    except Exception:
        traceback.print_exc()
    finally:
        sys.stdout = old_stdout
        sys.stderr = old_stderr
    text = output.getvalue()
    if activation_message:
        text = activation_message + "\n" + text
    return text


def render_page(message="", log=""):
    ensure_folders()
    entrada = list_files(converter.XML_FOLDER, "*.xml")
    salidas = list_files(converter.SALIDA_EXCEL_DIR, "*.xlsx")
    procesados = list_files(converter.PROCESADOS_DIR, "*.xml")
    errores = list_files(converter.ERRORES_DIR, "*.xml")

    def file_rows(files, kind):
        if not files:
            return '<div class="empty">Sin archivos</div>'
        rows = []
        for file in files[:12]:
            size_kb = max(1, file.stat().st_size // 1024)
            rows.append(
                "<li>"
                f"<span>{html.escape(file.name)}</span>"
                f"<small>{size_kb} KB</small>"
                f'<a href="/download?kind={kind}&name={quote(file.name)}">Descargar</a>'
                "</li>"
            )
        return "<ul>" + "".join(rows) + "</ul>"

    message_html = f'<div class="notice">{html.escape(message)}</div>' if message else ""
    log_html = f"<pre>{html.escape(log)}</pre>" if log else ""

    return f"""<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Consolidado XML a Excel</title>
  <style>
    :root {{
      color-scheme: dark;
      --bg: #0f1720;
      --panel: #151f2d;
      --panel-2: #1c2a3c;
      --text: #eef4ff;
      --muted: #9fb0c8;
      --line: #314156;
      --gold: #f5b544;
      --blue: #2f80ed;
      --red: #b54718;
      --green: #2e9d62;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: Arial, Helvetica, sans-serif;
      background: var(--bg);
      color: var(--text);
    }}
    header {{
      padding: 20px 28px;
      border-bottom: 1px solid var(--line);
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
    }}
    h1 {{ margin: 0; font-size: 24px; color: var(--gold); }}
    main {{ padding: 24px; max-width: 1180px; margin: 0 auto; }}
    .grid {{ display: grid; grid-template-columns: 1.1fr .9fr; gap: 18px; }}
    .panel {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 18px;
    }}
    .actions {{ display: grid; gap: 12px; }}
    button, .button {{
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 13px 16px;
      font-weight: 700;
      color: var(--text);
      background: var(--panel-2);
      cursor: pointer;
      text-align: center;
      text-decoration: none;
      font-size: 15px;
    }}
    button.primary {{ background: #183252; border-color: #365b89; }}
    button.danger {{ background: #7a2e12; border-color: var(--gold); }}
    button.light, .button.light {{ background: #f5f7fb; color: #101923; }}
    input[type=file], input[type=password] {{
      width: 100%;
      background: #0b1119;
      border: 1px dashed var(--line);
      border-radius: 6px;
      padding: 14px;
      color: var(--text);
    }}
    .notice {{
      margin-bottom: 16px;
      padding: 12px 14px;
      background: #183b2a;
      border: 1px solid var(--green);
      border-radius: 6px;
    }}
    h2 {{ font-size: 17px; margin: 0 0 12px; }}
    .cards {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; margin-top: 18px; }}
    ul {{ list-style: none; padding: 0; margin: 0; display: grid; gap: 8px; }}
    li {{
      display: grid;
      grid-template-columns: minmax(0, 1fr) 70px 80px;
      gap: 10px;
      align-items: center;
      padding: 9px 10px;
      background: #0b1119;
      border: 1px solid #243248;
      border-radius: 6px;
    }}
    li span {{ overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
    li small {{ color: var(--muted); text-align: right; }}
    li a {{ color: #8fc3ff; text-decoration: none; text-align: right; }}
    .empty {{ color: var(--muted); background: #0b1119; border-radius: 6px; padding: 12px; }}
    pre {{
      white-space: pre-wrap;
      overflow: auto;
      max-height: 300px;
      background: #05080d;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 12px;
      color: #cbd7e9;
    }}
    .brand {{
      border-top: 1px solid var(--line);
      margin-top: 16px;
      padding-top: 14px;
      color: var(--muted);
      line-height: 1.5;
    }}
    @media (max-width: 880px) {{
      .grid, .cards {{ grid-template-columns: 1fr; }}
      header {{ align-items: flex-start; flex-direction: column; }}
    }}
  </style>
</head>
<body>
  <header>
    <h1>CONSOLIDADO XML A EXCEL</h1>
    <div>App local privada</div>
  </header>
  <main>
    {message_html}
    <div class="grid">
      <section class="panel">
        <h2>Procesar XML</h2>
        <form class="actions" method="post" action="/upload" enctype="multipart/form-data">
          <input type="file" name="files" accept=".xml" multiple>
          <button class="primary" type="submit">CARGAR XML A ENTRADA</button>
        </form>
        <form class="actions" method="post" action="/process" style="margin-top:12px">
          <input name="clave" type="password" placeholder="Clave de instalacion si es la primera vez" autocomplete="off">
          <button class="primary" type="submit">EXTRAER INFORMACION A EXCEL</button>
        </form>
        <form class="actions" method="post" action="/clear" style="margin-top:12px">
          <button class="danger" type="submit">ELIMINAR XML Y PDF DESCARGADOS</button>
        </form>
        <div class="actions" style="margin-top:12px">
          <a class="button light" href="/open?folder=base">ABRIR CARPETA DEL PROGRAMA</a>
          <a class="button light" href="/open?folder=salida">ABRIR SALIDA EXCEL</a>
        </div>
        <div class="brand">
          <strong>Ing. Pablo Ronquillo</strong><br>
          Quito - Ecuador
        </div>
      </section>
      <section class="panel">
        <h2>Resultado</h2>
        {log_html or '<div class="empty">Listo.</div>'}
      </section>
    </div>
    <div class="cards">
      <section class="panel"><h2>Entrada XML ({len(entrada)})</h2>{file_rows(entrada, "entrada")}</section>
      <section class="panel"><h2>Salida Excel ({len(salidas)})</h2>{file_rows(salidas, "salida")}</section>
      <section class="panel"><h2>Procesados ({len(procesados)})</h2>{file_rows(procesados, "procesados")}</section>
      <section class="panel"><h2>Errores ({len(errores)})</h2>{file_rows(errores, "errores")}</section>
    </div>
  </main>
</body>
</html>"""


class AppHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.send_html(render_page())
            return
        if parsed.path == "/download":
            self.handle_download(parsed)
            return
        if parsed.path == "/open":
            self.handle_open(parsed)
            return
        self.send_error(404)

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == "/upload":
            self.handle_upload()
            return
        if parsed.path == "/process":
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length).decode("utf-8", errors="replace")
            clave = parse_qs(body).get("clave", [""])[0]
            log = run_converter(clave)
            self.send_html(render_page("Proceso terminado.", log))
            return
        if parsed.path == "/clear":
            self.handle_clear()
            return
        self.send_error(404)

    def send_html(self, body):
        data = body.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def handle_upload(self):
        ensure_folders()
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": self.headers.get("Content-Type"),
            },
        )
        items = form["files"] if "files" in form else []
        if not isinstance(items, list):
            items = [items]
        saved = 0
        for item in items:
            if not getattr(item, "filename", ""):
                continue
            filename = Path(item.filename).name
            if not filename.lower().endswith(".xml"):
                continue
            target = unique_path(converter.XML_FOLDER / filename)
            with target.open("wb") as fh:
                shutil.copyfileobj(item.file, fh)
            saved += 1
        self.send_html(render_page(f"XML cargados: {saved}"))

    def handle_download(self, parsed):
        query = parse_qs(parsed.query)
        kind = query.get("kind", [""])[0]
        name = unquote(query.get("name", [""])[0])
        folders = {
            "entrada": converter.XML_FOLDER,
            "salida": converter.SALIDA_EXCEL_DIR,
            "procesados": converter.PROCESADOS_DIR,
            "errores": converter.ERRORES_DIR,
        }
        if kind not in folders or not name:
            self.send_error(404)
            return
        try:
            path = safe_child(folders[kind], name)
        except ValueError:
            self.send_error(403)
            return
        if not path.is_file():
            self.send_error(404)
            return
        content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        data = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Disposition", f'attachment; filename="{path.name}"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def handle_open(self, parsed):
        query = parse_qs(parsed.query)
        folder = query.get("folder", ["base"])[0]
        target = converter.SALIDA_EXCEL_DIR if folder == "salida" else BASE_DIR
        open_in_explorer(target)
        self.send_html(render_page("Carpeta abierta en Windows."))

    def handle_clear(self):
        ensure_folders()
        count = 0
        for folder in (converter.XML_FOLDER, converter.PDF_ROOT_DIR):
            for pattern in ("*.xml", "*.pdf"):
                for file in folder.rglob(pattern):
                    try:
                        file.unlink()
                        count += 1
                    except OSError:
                        pass
        self.send_html(render_page(f"Archivos eliminados: {count}"))

    def log_message(self, fmt, *args):
        return


def main():
    ensure_folders()
    url = f"http://{HOST}:{PORT}"
    server = ThreadingHTTPServer((HOST, PORT), AppHandler)
    threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    print(f"App abierta en {url}", flush=True)
    print("Si el navegador no se abre solo, copia esa direccion en Chrome.", flush=True)
    print("Cierra esta ventana para apagar la app.", flush=True)
    server.serve_forever()


if __name__ == "__main__":
    main()
