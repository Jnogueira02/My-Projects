import textwrap
from mycanvas import MyCanvas


class TextObject2:
    x = 0
    y = 0
    text = ""
    text_list = []
    canvas = MyCanvas("")
    reverse = ""
    reverse_list = []
    font_size = 8
    line_char_count = 21
    spacing = 9
    reverse_font = "Helvetica"

    def __init__(self, x, y, canvas, text, reverse, reverse_font):
        self.x = x
        self.y = y
        self.canvas = canvas
        self.text = text
        self.reverse = reverse
        self.font_size, self.line_char_count, self.spacing = self.font_size()
        self.text_list = self.text_list()
        self.reverse_font = reverse_font
        self.reverse_list = self.reverse_list()

    def __str__(self):
        wrapped_string = textwrap.fill(self.text, self.line_char_count)
        dedented_string = textwrap.dedent(wrapped_string)
        formatted_string = dedented_string.replace("\n", "\n")
        return formatted_string

    def reversed(self):
        if self.reverse_font == "Courier-Bold":
            wrapped_string = textwrap.fill(self.reverse, 32)
        elif self.reverse_font == "Helvetica-Bold" or self.reverse_font == "Courier":
            wrapped_string = textwrap.fill(self.reverse, 33)
        else:
            wrapped_string = textwrap.fill(self.reverse, 35)
        dedented_string = textwrap.dedent(wrapped_string)
        formatted_string = dedented_string.replace("\n", "\n")
        return formatted_string

    def text_list(self):
        formatted_string = self.__str__()
        return formatted_string.split("\n")

    def reverse_list(self):
        formatted_string = self.reversed()
        return formatted_string.split("\n")

    def font_size(self):
        char_count = len(self.text)
        line_char_count = 22  # 21
        font_size = 8
        spacing = 9
        if char_count >= 130:
            None
        elif 110 <= char_count < 130:
            # font_size = 8
            line_char_count = 21
        elif 90 <= char_count < 110:
            font_size = 9
            line_char_count = 18
        elif 70 <= char_count < 90:
            font_size = 10
            line_char_count = 16
            spacing = 10
        elif 50 <= char_count < 70:
            font_size = 11
            line_char_count = 16
            spacing = 11
        else:
            font_size = 12
            line_char_count = 15
            spacing = 12

        if self.canvas.current_font == "Courier" or self.canvas.current_font == "Helvetica-Bold":
            font_size -= 1
            line_char_count -= 1
            spacing -= 1
        elif self.canvas.current_font == "Courier-Bold":
            font_size -= 2
            line_char_count -= 2
            spacing -= 2

        return font_size, line_char_count, spacing

    def print_string(self):
        for line in self.text_list:
            self.canvas.drawCentredString(self.x, self.y, line)
            # self.y -= 9
            self.y -= self.spacing

    def print_reverse(self):
        self.y -= 5
        self.canvas.setFont("Helvetica", 5)

        self.canvas.setFont(self.reverse_font, 5)

        for line in self.reverse_list:
            self.canvas.drawCentredString(self.x, self.y, line)
            self.y -= 5.5
