#!/usr/bin/env python3

import http.server
import io
import os
import base64
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE

SLIDE_W = 13.33
SLIDE_H = 7.5
MARGIN = 0.5
server_port = int(os.getenv('HTML2PPTX_PORT', '8080'))

# ── Color palette ──────────────────────────────────────────────────────────
BG       = RGBColor(0x0F, 0x17, 0x2A)   # dark navy
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT   = RGBColor(0x3B, 0x82, 0xF6)   # blue
ACCENT2  = RGBColor(0x10, 0xB9, 0x81)   # green
ACCENT3  = RGBColor(0xF5, 0x9E, 0x0B)   # amber
SUBTEXT  = RGBColor(0x94, 0xA3, 0xB8)   # slate

def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_title(slide, text, y=MARGIN, color=WHITE, size=36):
    tb = slide.shapes.add_textbox(Inches(MARGIN), Inches(y),
                                   Inches(SLIDE_W - 2*MARGIN), Inches(1.1))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = True
    p.font.color.rgb = color

def add_accent_bar(slide):
    """Thin blue bar under the title"""
    bar = slide.shapes.add_shape(
        1,  # rectangle
        Inches(MARGIN), Inches(1.45),
        Inches(SLIDE_W - 2*MARGIN), Inches(0.04)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = ACCENT
    bar.line.fill.background()

def add_bullets(slide, items, y=1.6, size=16):
    tb = slide.shapes.add_textbox(
        Inches(MARGIN), Inches(y),
        Inches(SLIDE_W - 2*MARGIN),
        Inches(SLIDE_H - y - MARGIN)
    )
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = '  •  ' + item
        p.font.size = Pt(size)
        p.font.color.rgb = RGBColor(0xCD, 0xD6, 0xF4)
        p.space_before = Pt(6)

def add_paragraph_text(slide, text, y=1.6):
    tb = slide.shapes.add_textbox(
        Inches(MARGIN), Inches(y),
        Inches(SLIDE_W - 2*MARGIN), Inches(1.5)
    )
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(16)
    p.font.color.rgb = SUBTEXT
    p.font.italic = True

def add_bar_chart(slide, chart_title, categories, series_name, values, y=1.7):
    chart_data = ChartData()
    chart_data.categories = categories
    chart_data.add_series(series_name, values)

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(MARGIN), Inches(y),
        Inches(SLIDE_W - 2*MARGIN), Inches(SLIDE_H - y - MARGIN),
        chart_data
    ).chart

    # Style the chart
    chart.has_title = True
    chart.chart_title.text_frame.text = chart_title
    chart.chart_title.text_frame.paragraphs[0].font.color.rgb = WHITE
    chart.chart_title.text_frame.paragraphs[0].font.size = Pt(14)

    plot = chart.plots[0]
    plot.has_data_labels = True
    plot.data_labels.font.color.rgb = WHITE
    plot.data_labels.font.size = Pt(11)

    # Color the bars
    colors = [ACCENT, ACCENT2, ACCENT3,
              RGBColor(0xEC, 0x48, 0x99), RGBColor(0x8B, 0x5C, 0xF6)]
    for i, point in enumerate(plot.series[0].points):
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = colors[i % len(colors)]

    # Style axes
    chart.category_axis.tick_labels.font.color.rgb = SUBTEXT
    chart.value_axis.tick_labels.font.color.rgb = SUBTEXT
    chart.category_axis.format.line.color.rgb = SUBTEXT
    chart.value_axis.format.line.color.rgb = SUBTEXT
    chart.plot_area.format.fill.background()

def add_pie_chart(slide, chart_title, categories, values, y=1.7):
    chart_data = ChartData()
    chart_data.categories = categories
    chart_data.add_series('', values)

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE,
        Inches(MARGIN), Inches(y),
        Inches(SLIDE_W - 2*MARGIN), Inches(SLIDE_H - y - MARGIN),
        chart_data
    ).chart

    chart.has_title = True
    chart.chart_title.text_frame.text = chart_title
    chart.chart_title.text_frame.paragraphs[0].font.color.rgb = WHITE
    chart.chart_title.text_frame.paragraphs[0].font.size = Pt(14)

    plot = chart.plots[0]
    plot.has_data_labels = True
    plot.data_labels.font.color.rgb = WHITE
    plot.data_labels.font.size = Pt(12)
    plot.data_labels.number_format = '0"%"'

    colors = [ACCENT, ACCENT2, ACCENT3,
              RGBColor(0xEC, 0x48, 0x99), RGBColor(0x8B, 0x5C, 0xF6)]
    for i, point in enumerate(plot.series[0].points):
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = colors[i % len(colors)]

def add_two_column(slide, left_heading, left_items, right_heading, right_items):
    col_w = (SLIDE_W - 2*MARGIN - 0.3) / 2

    for col_x, heading, items in [
        (MARGIN, left_heading, left_items),
        (MARGIN + col_w + 0.3, right_heading, right_items)
    ]:
        # Column background card
        card = slide.shapes.add_shape(
            1, Inches(col_x), Inches(1.7),
            Inches(col_w), Inches(SLIDE_H - 1.7 - MARGIN)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(0x1E, 0x29, 0x3B)
        card.line.color.rgb = ACCENT

        # Heading
        tb = slide.shapes.add_textbox(
            Inches(col_x + 0.2), Inches(1.85),
            Inches(col_w - 0.4), Inches(0.6)
        )
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = heading
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = ACCENT

        # Bullets
        tb2 = slide.shapes.add_textbox(
            Inches(col_x + 0.2), Inches(2.55),
            Inches(col_w - 0.4), Inches(SLIDE_H - 2.55 - MARGIN - 0.3)
        )
        tf2 = tb2.text_frame
        tf2.word_wrap = True
        tf2.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        for i, item in enumerate(items):
            p = tf2.paragraphs[0] if i == 0 else tf2.add_paragraph()
            p.text = '• ' + item
            p.font.size = Pt(15)
            p.font.color.rgb = RGBColor(0xCD, 0xD6, 0xF4)
            p.space_before = Pt(5)

def html_to_pptx_bytes(html_string):
    soup = BeautifulSoup(html_string, 'html.parser')
    slide_els = soup.select('section.slide')

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    blank = prs.slide_layouts[6]

    for el in slide_els:
        slide = prs.slides.add_slide(blank)
        set_bg(slide, BG)

        title_el = el.find('h1')
        title_text = title_el.get_text().strip() if title_el else 'Untitled'
        add_title(slide, title_text)
        add_accent_bar(slide)

        chart_type = el.get('data-chart', '').strip()

        if chart_type == 'bar':
            cats = [c.strip() for c in el.get('data-categories', '').split(',')]
            vals = [float(v.strip()) for v in el.get('data-values', '').split(',')]
            series_name = el.get('data-series-name', '')
            chart_title = el.get('data-chart-title', '')
            p_el = el.find('p')
            if p_el:
                add_paragraph_text(slide, p_el.get_text().strip())
                add_bar_chart(slide, chart_title, cats, series_name, vals, y=2.3)
            else:
                add_bar_chart(slide, chart_title, cats, series_name, vals)

        elif chart_type == 'pie':
            cats = [c.strip() for c in el.get('data-categories', '').split(',')]
            vals = [float(v.strip()) for v in el.get('data-values', '').split(',')]
            chart_title = el.get('data-chart-title', '')
            p_el = el.find('p')
            if p_el:
                add_paragraph_text(slide, p_el.get_text().strip())
                add_pie_chart(slide, chart_title, cats, vals, y=2.3)
            else:
                add_pie_chart(slide, chart_title, cats, vals)

        elif el.get('data-layout') == 'two-column':
            left = el.find(class_='col-left')
            right = el.find(class_='col-right')
            lh = left.find('h2').get_text().strip() if left and left.find('h2') else ''
            rh = right.find('h2').get_text().strip() if right and right.find('h2') else ''
            li = [i.get_text().strip() for i in left.find_all('li')] if left else []
            ri = [i.get_text().strip() for i in right.find_all('li')] if right else []
            add_two_column(slide, lh, li, rh, ri)

        else:
            # Default: bullets
            items = [li.get_text().strip() for li in el.find_all('li')]
            if items:
                add_bullets(slide, items)

    if not slide_els:
        slide = prs.slides.add_slide(blank)
        set_bg(slide, BG)
        tb = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(11), Inches(1))
        tb.text_frame.paragraphs[0].text = 'No slides detected'

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


class Handler(http.server.BaseHTTPRequestHandler):
    def do_POST(self):
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length).decode('utf-8').strip()

        try:
            html_string = base64.b64decode(body).decode('utf-8')
        except Exception:
            html_string = body

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
