import http.server
import socketserver
import os
import json

PORT = 8000
WORKBOOK_FILE = 'Planilha Cris.xlsx'
CATEGORIES_FILE = 'categories.json'
FORMATTING_FILE = 'formatting.json'
CONDOMINIOS_FILE = 'condominios.json'

class CustomHandler(http.server.SimpleHTTPRequestHandler):

    def _send_json(self, code, data):
        body = json.dumps(data).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, path, content_type):
        with open(path, 'rb') as f:
            data = f.read()
        self.send_response(200)
        self.send_header('Content-Type', content_type)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    # ── POST ──────────────────────────────────────────────────────
    def do_POST(self):
        length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(length)

        try:
            if self.path == '/save_workbook':
                with open(WORKBOOK_FILE, 'wb') as f:
                    f.write(body)
                self._send_json(200, {'status': 'ok'})

            elif self.path == '/save_categories':
                with open(CATEGORIES_FILE, 'w', encoding='utf-8') as f:
                    f.write(body.decode('utf-8'))
                self._send_json(200, {'status': 'ok'})

            elif self.path == '/save_formatting':
                with open(FORMATTING_FILE, 'w', encoding='utf-8') as f:
                    f.write(body.decode('utf-8'))
                self._send_json(200, {'status': 'ok'})

            elif self.path == '/save_condominios':
                with open(CONDOMINIOS_FILE, 'w', encoding='utf-8') as f:
                    f.write(body.decode('utf-8'))
                self._send_json(200, {'status': 'ok'})

            else:
                self._send_json(404, {'error': 'Not found'})

        except Exception as e:
            self._send_json(500, {'error': str(e)})

    # ── OPTIONS (CORS preflight) ───────────────────────────────────
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    # ── GET ───────────────────────────────────────────────────────
    def do_GET(self):
        try:
            if self.path == '/status':
                self._send_json(200, {'status': 'ok', 'server': 'Karpa v2'})

            elif self.path == '/load_workbook':
                if os.path.exists(WORKBOOK_FILE):
                    self._send_file(WORKBOOK_FILE,
                        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                else:
                    self._send_json(404, {'error': 'Workbook not found'})

            elif self.path == '/load_categories':
                if os.path.exists(CATEGORIES_FILE):
                    with open(CATEGORIES_FILE, 'r', encoding='utf-8') as f:
                        self.send_response(200)
                        self.send_header('Content-Type', 'application/json')
                        self.send_header('Access-Control-Allow-Origin', '*')
                        self.end_headers()
                        self.wfile.write(f.read().encode('utf-8'))
                else:
                    self._send_json(200, {})

            elif self.path == '/load_formatting':
                if os.path.exists(FORMATTING_FILE):
                    with open(FORMATTING_FILE, 'r', encoding='utf-8') as f:
                        self.send_response(200)
                        self.send_header('Content-Type', 'application/json')
                        self.send_header('Access-Control-Allow-Origin', '*')
                        self.end_headers()
                        self.wfile.write(f.read().encode('utf-8'))
                else:
                    self._send_json(200, {})

            elif self.path == '/load_condominios':
                if os.path.exists(CONDOMINIOS_FILE):
                    with open(CONDOMINIOS_FILE, 'r', encoding='utf-8') as f:
                        self.send_response(200)
                        self.send_header('Content-Type', 'application/json')
                        self.send_header('Access-Control-Allow-Origin', '*')
                        self.end_headers()
                        self.wfile.write(f.read().encode('utf-8'))
                else:
                    self._send_json(200, [])

            else:
                super().do_GET()

        except Exception as e:
            self._send_json(500, {'error': str(e)})

    def log_message(self, fmt, *args):
        # Silencia logs de GET de arquivos estáticos comuns
        if args and isinstance(args[0], str) and any(ext in args[0] for ext in ['.js', '.css', '.html', '.ico', '.png']):
            return
        super().log_message(fmt, *args)


# ── Server bootstrap ──────────────────────────────────────────────
socketserver.TCPServer.allow_reuse_address = True

with socketserver.TCPServer(('', PORT), CustomHandler) as httpd:
    print(f'\n✅  Karpa Server iniciado na porta {PORT}')
    print(f'   Acesse: http://localhost:{PORT}\n')
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print('\n⛔  Servidor encerrado.')
        httpd.server_close()
