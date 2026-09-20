import os
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
from PIL import Image, ImageDraw
import pdfplumber

pdf_path = os.path.join(os.getcwd(), 'sample_layout.pdf')

c = canvas.Canvas(pdf_path, pagesize=letter)
c.setFont('Helvetica-Bold', 14)
c.drawString(70, 740, 'Name')
c.drawString(220, 740, 'Age')
c.drawString(70, 710, 'Alice')
c.drawString(220, 710, '28')
c.drawString(70, 680, 'Bob')
c.drawString(220, 680, '31')

img = Image.new('RGB', (120, 60), color='white')
d = ImageDraw.Draw(img)
d.rectangle((10, 10, 110, 50), fill='lightblue')
d.text((20, 20), 'LOGO')
img_buf = BytesIO()
img.save(img_buf, format='PNG')
img_buf.seek(0)
c.drawImage(ImageReader(img_buf), 400, 690, width=120, height=60)
c.save()

print('CREATED', pdf_path)

with pdfplumber.open(pdf_path) as pdf:
    page = pdf.pages[0]
    print('TABLES', page.extract_tables()[:1])
    print('WORDS_COUNT', len(page.extract_words()))
    print('IMAGE_COUNT', len(page.images))
    print('IMAGE_KEYS', page.images[0].keys() if page.images else [])
    print('IMAGE_STREAM_LEN', len(page.images[0]['stream']) if page.images else 0)
