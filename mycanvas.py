# Import the Canvas class from reportlab (for creating and editing PDFs)
from reportlab.pdfgen import canvas

# Add a current_font attribute to the canvas class
class MyCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.current_font = None

# Set the current_font for the canvas object
    def setFont(self, fontname, fontsize):
        self.current_font = fontname
        super().setFont(fontname, fontsize)
