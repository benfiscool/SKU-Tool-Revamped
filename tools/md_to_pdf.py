"""
Simple Markdown to PDF converter using ReportLab.
This script converts USER_GUIDE.md to USER_GUIDE.pdf.

Usage:
  python tools/md_to_pdf.py

It requires reportlab to be installed: pip install reportlab
"""
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer
from reportlab.lib.units import inch
import markdown
import os

INPUT_MD = os.path.join(os.path.dirname(__file__), '..', 'USER_GUIDE.md')
OUTPUT_PDF = os.path.join(os.path.dirname(__file__), '..', 'USER_GUIDE.pdf')


def md_to_pdf(md_path, pdf_path):
    with open(md_path, 'r', encoding='utf-8') as f:
        md = f.read()
    # Convert markdown to simple HTML
    html = markdown.markdown(md)

    # Build a simple PDF using ReportLab's Platypus
    doc = SimpleDocTemplate(pdf_path, pagesize=letter,
                            rightMargin=72, leftMargin=72,
                            topMargin=72, bottomMargin=72)
    styles = getSampleStyleSheet()
    story = []

    # Split HTML by paragraphs
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, 'html.parser')
    for elem in soup.find_all(['h1','h2','h3','p','ul','ol','pre']):
        text = str(elem)
        style = styles['BodyText']
        if elem.name == 'h1':
            style = styles['Title']
        elif elem.name == 'h2':
            style = styles['Heading2']
        elif elem.name == 'h3':
            style = styles['Heading3']
        elif elem.name in ('ul','ol'):
            # convert list items
            for li in elem.find_all('li'):
                story.append(Paragraph(f'• {li.get_text()}', styles['BodyText']))
            story.append(Spacer(1, 0.1*inch))
            continue
        elif elem.name == 'pre':
            story.append(Paragraph(elem.get_text().replace('\n','<br/>'), styles['Code']))
            continue

        # Clean HTML tags for paragraph
        from html import unescape
        clean_text = unescape(elem.get_text())
        story.append(Paragraph(clean_text.replace('\n', '<br/>'), style))
        story.append(Spacer(1, 0.1*inch))

    doc.build(story)


if __name__ == '__main__':
    try:
        md_to_pdf(INPUT_MD, OUTPUT_PDF)
        print(f'PDF generated: {OUTPUT_PDF}')
    except Exception as e:
        print('Failed to generate PDF:', e)
        print('Make sure reportlab and beautifulsoup4 are installed: pip install reportlab beautifulsoup4 markdown')
