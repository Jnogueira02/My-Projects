# Import modules:

# MyCanvas for editing PDFs
# PIL for handling images
# NumPy for scaling arrays (used for adjusting image dimensions)
# Requests and BytesIO for accessing photos from the web
# Time for timing
from mycanvas import MyCanvas
from PIL import Image
import numpy as np
import requests
from io import BytesIO

import time


# Scale image dimensions to fit into PDF box
def scale_np_array(array_data, limit):
    scaling_factor = limit / array_data[0]
    scaled_np_array = array_data * scaling_factor

    if scaled_np_array[1] > limit:
        scaling_factor = limit / scaled_np_array[1]
        scaled_np_array *= scaling_factor

    return scaled_np_array


class FadedImage2:
    x = 0
    y = 0
    canvas = MyCanvas("")
    image_url = ""
    count = 1

    def __init__(self, x, y, canvas, image_url, count):
        self.x = x
        self.y = y
        self.canvas = canvas
        self.image_url = image_url
        self.count = count

    def print_pic_url(self):
        t1 = time.time()

        my_bool = True

        response = requests.get(self.image_url)
        image = Image.open(BytesIO(response.content))

        if my_bool:
            width, height = image.size

            image.save(f"MyPica44edtfbdv5yygr{self.count}.jpg")
            image.close()

            my_dimensions = np.array([width, height])
            my_scaled_dimensions = scale_np_array(my_dimensions, 104)

            self.x = self.x - (my_scaled_dimensions[0] / 2)
            self.y = self.y - (my_scaled_dimensions[1] / 2)

            self.canvas.setFillAlpha(0.5)
            self.canvas.drawImage(f"MyPica44edtfbdv5yygr{self.count}.jpg", self.x, self.y + 26.8,
                                  my_scaled_dimensions[0], my_scaled_dimensions[1])
            self.canvas.setFillAlpha(1)

            print(f"print_pic_2 time: {time.time() - t1}")
            return f"MyPica44edtfbdv5yygr{self.count}.jpg"  ##Need to change this??

    def print_pic_downloaded(self, opacity):
        t1 = time.time()

        image = Image.open(self.image_url)

        width, height = image.size

        image.close()

        my_dimensions = np.array([width, height])
        my_scaled_dimensions = scale_np_array(my_dimensions, 104)

        self.x = self.x - (my_scaled_dimensions[0] / 2)
        self.y = self.y - (my_scaled_dimensions[1] / 2)

        self.canvas.setFillAlpha(int(opacity) / 100)

        self.canvas.drawImage(self.image_url, self.x - 3.15, self.y + 26.8, my_scaled_dimensions[0],
                              my_scaled_dimensions[1])
        self.canvas.setFillAlpha(1)

        print(f"print_pic_2 time: {time.time() - t1}")
        return self.image_url
