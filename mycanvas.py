from reportlab.pdfgen import canvas


class MyCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.current_font = None

    def setFont(self, fontname, fontsize):
        self.current_font = fontname
        super().setFont(fontname, fontsize)
