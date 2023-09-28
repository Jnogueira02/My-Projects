import os
import PySimpleGUI as sg
import pandas as pd
import pandas.errors
from mycanvas import MyCanvas
from textobject2 import TextObject2
from fadedimage2 import FadedImage2
import math

import time

sg.theme("DarkPurple")

layout = [
    [[sg.Text("Choose an Excel file: "), sg.FileBrowse(key="file_name")]],
    [sg.Text("Output file name (PDF): "), sg.InputText(key="output")],
    [sg.Text("Choose an opacity (%): "), sg.InputText(default_text="40", key="opacity")],
    [sg.Text("Choose a serial font: ", size=(20,1)), sg.Combo(["Courier", "Helvetica", "Times-Roman", "Helvetica-Bold", "Courier-Bold" ], default_value="Courier", key="font")],
    [sg.Text("Choose a description font: ", size=(20, 1)),
     sg.Combo(["Courier", "Helvetica", "Times-Roman", "Helvetica-Bold", "Courier-Bold"], default_value="Courier-Bold",
              key="font2")],
    [sg.Text("Choose a reverse font: ", size=(20, 1)),
     sg.Combo(["Courier", "Helvetica", "Times-Roman", "Helvetica-Bold", "Courier-Bold"], default_value="Helvetica",
              key="font3")],
    [sg.Submit(), sg.Button("Clear"), sg.Exit()]
]

window = sg.Window("Insertion", layout, size=(600, 200))


def clear_input():
    for key in values:
        if key != "file_name":
            window[key]("")
    return None


count = 1
my_bool = True

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event == "Clear":
        clear_input()
    if event == "Submit":
        t1 = time.time()
        list_values = list(values.values())
        print(list_values)
        print("\n*******************************\n")

        file_name = list_values[0]
        try:
            df = pd.read_excel(file_name, sheet_name=0)
        except pandas.errors.ParserError:
            window["output"].update("Error: Invalid Excel File")
            continue
        except FileNotFoundError:
            window["output"].update("Error: No such Excel file")
            continue

        try:
            print(df.columns)
            print(df.shape)

            if df.iloc[:, 2].isnull().any():
                cleaned_data = df.dropna(subset=df.columns[[0, 1]])
            else:
                cleaned_data = df.dropna(subset=df.columns[[0, 1, 2]])

        except IndexError:
            window["output"].update("Error: expected 2 or 3 columns")
            continue
        except KeyError:
            window["output"].update("Error: Expected description/serial column/s not found") #Update later to specific column not found
            continue

        try:
            print(cleaned_data)
        except NameError:
            continue
        print("\n*******************************\n")

        selected_values = cleaned_data.values
        print(selected_values[:5])
        print("\n*******************************\n")

        empty_count = (df.isna().all(axis=1)).sum() - 1

        font = list_values[-3]
        font2 = list_values[-2]
        font3 = list_values[-1]

        opacity = list_values[2]

        output_name = list_values[1]
        canvas = MyCanvas(f"{output_name}.pdf", pagesize=(612.0, 792.0))

        font_size = 11
        if font == "Courier":
            font_size = 10

        try:
            canvas.setFont(font3, font_size)
        except KeyError:
            window["output"].update("Error: Invalid Reverse Font")
            continue
        try:
            canvas.setFont(font2, font_size)
        except KeyError:
            window["output"].update("Error: Invalid Description Font")
            continue
        try:
            canvas.setFont(font, font_size)
        except KeyError:
            window["output"].update("Error: Invalid Serial Font")
            continue

        count = 1
        count2 = 1

        end = len(selected_values)

        nums = math.ceil(end / 24)
        print(nums)

        if count == 1:
            canvas.drawImage("Boxed Flip Sheet.jpeg", 41.11, 47, 544.2234, 704.2966)

        text_object_list = []
        faded_object_list = []
        my_pic_list = []

        start = 1

        errored_rows = []

        my_bool_2 = True
        my_bool_3 = True

        for num in range(nums):

            starting_x = 99.11

            for j in range(4):
                starting_y = 665
                starting_y_2 = 651
                for i in range(6):
                    # If start > end
                    if start > end:
                        my_bool_2 = False
                    if not my_bool_2:
                        break
                    image_url = selected_values[count - 1][2]
                    term = selected_values[count - 1][1]
                    text = selected_values[count - 1][0]

                    if isinstance(term, float):
                        term = ""

                    reverse = ""

                    if len(term) > 15:
                        my_list = term.split()
                        term = my_list[0]
                        reverse = " ".join(my_list[2:]) #######

                    if isinstance(text, float):
                        text = ""

                    # If the description > 150 characters
                    if len(text) > 150:
                        text = text[:150]
                        errored_rows.append(count + empty_count)
                        window["output"].update(f"Row/s {errored_rows} exceed/s 150 characters")

                    if not isinstance(image_url, float):
                        faded_object_list.append(FadedImage2(612.5 - starting_x, starting_y, canvas, image_url, count))
                    else:
                        faded_object_list.append(0.0)
                    text_object_list.append(TextObject2(612.5 - starting_x, starting_y + 68, canvas, text, reverse, font3))
                    canvas.drawCentredString(starting_x, starting_y_2, term)
                    start += 1
                    # starting_y -= 116.85
                    starting_y -= 116.65
                    starting_y_2 -= 116.65
                    count += 1
                # starting_x += 114
                starting_x += 114.7
            if count % 24 == 1 or count >= len(text_object_list):
                canvas.showPage()
                # canvas.drawImage("Reversed Flip Sheet.jpeg", 26.51, 46, 544.2234, 704.2966)
                # canvas.drawImage("Reversed Flip Sheet.jpeg", 24.51, 46, 544.2234, 704.2966)
                # canvas.drawImage("Reversed Flip Sheet.jpeg", 24.11, 46, 544.2234, 704.2966)
                # canvas.drawImage("Reversed Flip Sheet.jpeg", 23.71, 46, 544.2234, 704.2966)
                canvas.drawImage("Reversed Flip Sheet.jpeg", 23.11, 46, 544.2234, 704.2966)
            for j in range(4):
                for i in range(6):
                    if count2 > len(text_object_list):
                        my_bool_3 = False
                    if not my_bool_3:
                        break

                    current_image = faded_object_list[count2 - 1]
                    if not isinstance(current_image, float):
                        try:
                            if "https" in current_image.image_url:
                                my_pic_list.append(current_image.print_pic_url())
                            else:
                                my_pic_list.append(current_image.print_pic_downloaded(opacity))
                        except FileNotFoundError as e:
                            window["output"].update(e)
                        except OSError as e:
                            window["output"].update(e)
                        except ValueError:
                            window["opacity"].update(f"Invalid Opacity: {opacity}")

                    current_text_object = text_object_list[count2 - 1]
                    canvas.setFont(font2, current_text_object.font_size)
                    current_text_object.print_string()
                    current_text_object.print_reverse()
                    count2 += 1
            if count % 24 == 1:
                canvas.showPage()
                canvas.drawImage("Boxed Flip Sheet.jpeg", 41.11, 47, 544.2234, 704.2966)
                canvas.setFont(font, 12)
        try:
            canvas.save()
            print("Done")
            print(f"total time: {time.time() - t1}")
        except PermissionError:
            window["output"].update(f"Error: Output file must not be open while saving")
        except OSError:
            window["output"].update(f"Error: Invalid Output Filename")
        finally:
            for name in my_pic_list:
                # print(name, type(name))
                if isinstance(name, str):
                    if "MyPica44edtfbdv5yygr" in name:
                        os.remove(name)
            print(f"**{len(my_pic_list)}**")
            print("Done")
            print(f"total time: {time.time()-t1}")

window.close()
