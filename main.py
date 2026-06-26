import http.server
import io
import os
import base64
import json
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE

SLIDE_WIDTH_INCHES = 13.33
SLIDE_HEIGHT_INCHES = 7.5
MARGIN = 0.5
server_port = int(os.getenv('HTML2PPTX_PORT', '8080'))

def html_to_pptx_bytes(html_string):
    soup = BeautifulSoup(html_string, 'html.parser')
    slide_els = soup.select('.slide, section, [data-slide]')

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_WIDTH_INCHES)
    prs.slide_height = Inches(SLIDE_HEIGHT_INCHES)
    blank_layout = prs.slide_layouts[6]

    for el in slide_els:
        slide = prs.slides.add_slide(blank_layout)
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(0x1E, 0x1E, 0x2E)

        title_el = el.find(['h1', 'h2'])
        y_offset = MARGIN
        if title_el:
            txBox = slide.shapes.add_textbox(
                Inches(MARGIN), Inches(y_offset),
                Inches(SLIDE_WIDTH_INCHES - 2*MARGIN), Inches(1.2)
            )
            tf = txBox.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = title_el.get_text().strip()
            p.font.size = Pt(36)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            y_offset += 1.4

        sub_el = el.find('h3') or el.find(class_='subtitle')
        if sub_el:
            txBox = slide.shapes.add_textbox(
                Inches(MARGIN), Inches(y_offset),
                Inches(SLIDE_WIDTH_INCHES - 2*MARGIN), Inches(0.6)
            )
            tf = txBox.text_frame
            p = tf.paragraphs[0]
            p.text = sub_el.get_text().strip()
            p.font.size = Pt(20)
            p.font.italic = True
            p.font.color.rgb = RGBColor(0x89, 0xB4, 0xFA)
            y_offset += 0.8

        list_items = el.find_all('li')
        if list_items:
            txBox = slide.shapes.add_textbox(
                Inches(MARGIN), Inches(y_offset),
                Inches(SLIDE_WIDTH_INCHES - 2*MARGIN),
                Inches(SLIDE_HEIGHT_INCHES - y_offset - MARGIN)
            )
            tf = txBox.text_frame
            tf.word_wrap = True
            tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            for i, li in enumerate(list_items):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.text = '• ' + li.get_text().strip()
                p.font.size = Pt(18)
                p.font.color.rgb = RGBColor(0xCD, 0xD6, 0xF4)
        else:
            paragraphs = el.find_all('p')
            if paragraphs:
                body = '\n'.join(p.get_text().strip() for p in paragraphs)
                txBox = slide.shapes.add_textbox(
                    Inches(MARGIN), Inches(y_offset),
                    Inches(SLIDE_WIDTH_INCHES - 2*MARGIN),
                    Inches(SLIDE_HEIGHT_INCHES - y_offset - MARGIN)
                )
                tf = txBox.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = body
                p.font.size = Pt(16)
                p.font.color.rgb = RGBColor(0xCD, 0xD6, 0xF4)

    if len(slide_els) == 0:
        slide = prs.slides.add_slide(blank_layout)
        txBox = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(11), Inches(1))
        tf = txBox.text_frame
        tf.paragraphs[0].text = 'No slides detected — check your HTML selectors'

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


class Handler(http.server.BaseHTTPRequestHandler):
    def do_POST(self):
        content_type = self.headers.get('Content-Type', '')
        content_length = int(self.headers['Content-Length'])
        body = self.rfile.read(content_length)
    
        html_string = ''
    
        if 'multipart/form-data' in content_type:
            # Extract boundary
            boundary = None
            for part in content_type.split(';'):
                part = part.strip()
                if part.startswith('boundary='):
                    boundary = part[9:].strip()
                    break
    
            if boundary:
                boundary_bytes = ('--' + boundary).encode()
                parts = body.split(boundary_bytes)
                for part in parts:
                    if b'Content-Disposition' in part:
                        # Split headers from body
                        if b'\r\n\r\n' in part:
                            headers_raw, content = part.split(b'\r\n\r\n', 1)
                            # Strip trailing boundary markers
                            content = content.rstrip(b'\r\n--')
                            html_string = content.decode('utf-8')
                            break  # Take first file field found
    
        elif 'application/json' in content_type:
            import json, base64
            data = json.loads(body.decode('utf-8'))
            raw = data.get('html', '')
            try:
                html_string = base64.b64decode(raw).decode('utf-8')
            except Exception:
                html_string = raw
    
        else:
            # Raw body fallback
            html_string = body.decode('utf-8')
    
        pptx_bytes = html_to_pptx_bytes(html_string)
    
        self.send_response(200)
        self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
        self.send_header('Content-Disposition', 'attachment; filename="presentation.pptx"')
        self.send_header('Content-Length', str(len(pptx_bytes)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(pptx_bytes)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def log_message(self, format, *args):
        pass


httpd = http.server.HTTPServer(('', server_port), Handler)
print(f'html2pptx serving on port {server_port}')
httpd.serve_forever()
