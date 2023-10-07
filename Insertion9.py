# Import modules:

# OS for saving files
# PySimpleGUI for GUI
# Pandas for Excel
# MyCanvas for creating PDFs
# TextObject2 for formatting text
# Math for ceiling function
# Time for timing

import os
import PySimpleGUI as sg
import pandas as pd
import pandas.errors
from mycanvas import MyCanvas
from textobject2 import TextObject2
from fadedimage2 import FadedImage2
import math

import time

# Set the theme for the GUI
sg.theme("DarkPurple")

# Set the layout for the GUI:

# File Browser for accessing the Excel file
# Input Text for choosing an output PDF name
# Combo Box for choosing a font for the serial number
# Combo Box for choosing a font for the description
# Combo Box for choosing a font for the reverse description
# A Submit Button for executing the program
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

# Set the Window size and name
window = sg.Window("Insertion", layout, size=(600, 200))


# Clear input from all boxes except for the Excel file name
def clear_input():
    for key in values:
        if key != "file_name":
            window[key]("")
    return None


# Initialize count and boolean variables
count = 1
my_bool = True

# Create infinite loop for the program. Halts only when exit is clicked.
while True:
    # Event is the button that has been clicked. Values are the current data fields of the GUI
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event == "Clear":
        clear_input()
    if event == "Submit":
        # Start the timer
        t1 = time.time()
        # Put the values of the data fields into a list
        list_values = list(values.values())
        # Print (to standard output) the values, as well as a string of asterisks to signify the beginning of program execution
        print(list_values)
        print("\n*******************************\n")
        # Retrieve the file name from the list of values
        file_name = list_values[0]
        # Try to read the file into a pandas DataFrame. If an error is encountered, skip the rest of execution.
        try:
            df = pd.read_excel(file_name, sheet_name=0)
        except pandas.errors.ParserError:
            window["output"].update("Error: Invalid Excel File")
            continue
        except FileNotFoundError:
            window["output"].update("Error: No such Excel file")
            continue
        # Check that the DataFrame has at least 2 columns. If an error is encountered, skip the rest of execution.
        try:
            # Print the columns and shape of the DataFrame
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
            window["output"].update("Error: Expected description/serial column/s not found")
            continue

        # Try to print our DataFrame.  If an error is encountered, skip the rest of execution.
        try:
            print(cleaned_data)
        except NameError:
            continue
        # Print the first couple of values for the Excel data
        print("\n*******************************\n")

        selected_values = cleaned_data.values
        print(selected_values[:5])
        print("\n*******************************\n")

        # Count the whitespace from the top of the Excel file
        empty_count = (df.isna().all(axis=1)).sum() - 1

        # Retrieve the various font choices, opacity %, and output filename entered into the GUI
        font = list_values[-3]
        font2 = list_values[-2]
        font3 = list_values[-1]

        opacity = list_values[2]

        output_name = list_values[1]

        # Create Canvas object to create and edit our PDF
        canvas = MyCanvas(f"{output_name}.pdf", pagesize=(612.0, 792.0))
        # Choose the default font size for the Canvas
        font_size = 11
        if font == "Courier":
            font_size = 10

        # Try to set the fonts for our Canvas object.  If an error is encountered (invalid font size), skip the rest of execution.
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

        # Initialize our count variables to 1
        count = 1
        count2 = 1

        # Calculate the number of rows in our Excel file. This corresponds to the number of PDF pages we will need (each page will hold 24 rows).
        end = len(selected_values)
        nums = math.ceil(end / 24)
        print(nums)

        # Draw our Boxed Flip Sheet image on our PDF
        if count == 1:
            canvas.drawImage("Boxed Flip Sheet.jpeg", 41.11, 47, 544.2234, 704.2966)

        # Initialize our lists for holding our TextOnjects, FadedImages, and pictures
        text_object_list = []
        faded_object_list = []
        my_pic_list = []

        # Initialize start variable
        start = 1

        # Create list to hold rows that may throw an error in the future. Used for debugging.
        errored_rows = []

        # Create boolean variables
        my_bool_2 = True
        my_bool_3 = True

        # For loop for adding and editing pages to PDF
        for num in range(nums):
            # Starting x-coordinates for writing text on the PDF
            starting_x = 99.11
            # Iterate through columns
            for j in range(4):
                # Starting y-coordinates for writing text on the PDF
                starting_y = 665
                starting_y_2 = 651
                # Iterate through rows
                for i in range(6):
                    # If start > end, break out of loop
                    if start > end:
                        break
                    # Find the image_url/saved picture name, serial number, and description of the current row (by using our count variable)
                    image_url = selected_values[count - 1][2]
                    term = selected_values[count - 1][1]
                    text = selected_values[count - 1][0]
                    # If the term is NULL, make it equal to the empty string, ""
                    if isinstance(term, float):
                        term = ""
                    # Initialize the reverse description
                    reverse = ""

                    # Isolate the reverse description from the serial number (before the first space = serial number, after second space = reverse description)
                    if len(term) > 15:
                        my_list = term.split()
                        term = my_list[0]
                        reverse = " ".join(my_list[2:]) #######

                    # If the description is NULL, make it equal to the empty string, ""
                    if isinstance(text, float):
                        text = ""

                    # If the description > 150 characters, select only the first 150. A description greater than 50 will result in formatting issues.
                    if len(text) > 150:
                        text = text[:150]
                        errored_rows.append(count + empty_count)
                        window["output"].update(f"Row/s {errored_rows} exceed/s 150 characters")

                    # If the image_url/saved picture name is not NULL, create a FadedImage object with the relevant x and y coordinates for the PDF,
                    # and add it to our FadedImage Object list. Otherwise, add a filler value to the list.
                    if not isinstance(image_url, float):
                        faded_object_list.append(FadedImage2(612.5 - starting_x, starting_y, canvas, image_url, count))
                    else:
                        faded_object_list.append(0.0)
                    # Create a TextObject with the relevant x and y coordinates for the PDF (used for the description).
                    text_object_list.append(TextObject2(612.5 - starting_x, starting_y + 68, canvas, text, reverse, font3))
                    # Draw the serial number on the PDF
                    canvas.drawCentredString(starting_x, starting_y_2, term)
                    # Update start and count variables, as well as y-coordinates
                    start += 1
                    starting_y -= 116.65
                    starting_y_2 -= 116.65
                    count += 1
                # Update x-coordinate
                starting_x += 114.7
            # If a page is complete (a multiple of 24 variables or number of Excel rows < 24).
            if count % 24 == 1 or count >= len(text_object_list):
                # Save the current page and make a back page to hold the descriptions
                canvas.showPage()
                canvas.drawImage("Reversed Flip Sheet.jpeg", 23.11, 46, 544.2234, 704.2966)
            # Iterate through columns
            for j in range(4):
                # Iterate through rows
                for i in range(6):
                    # If our second variable is greater than the length of our Excel rows, break from the loop
                    if count2 > len(text_object_list):
                        break

                    # Find the image in our FadedImage object list relevant to our second count variable
                    current_image = faded_object_list[count2 - 1]
                    # If the image is not NULL, print it. Catch any errors that may occur (invalid file name, invalid opacity input, etc.).
                    if not isinstance(current_image, float):
                        # If the current image object contains "https", call the function to print a photo from the web.
                        # Otherwise, call the function to print a downloaded photo
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

                    # Select the current text object from our TextObject list, set the appropriate font, and print the description and reverse description.
                    current_text_object = text_object_list[count2 - 1]
                    canvas.setFont(font2, current_text_object.font_size)
                    current_text_object.print_string()
                    current_text_object.print_reverse()
                    # Update the count variable
                    count2 += 1
            # If a page is complete (a multiple of 24 variables), save the current page and make a new front page to hold the serial numbers.
            # Also, reset the font size to default (size 12)
            if count % 24 == 1:
                canvas.showPage()
                canvas.drawImage("Boxed Flip Sheet.jpeg", 41.11, 47, 544.2234, 704.2966)
                canvas.setFont(font, 12)
        # Try to save the PDF file in the current directory, and print the time it took for the program to run to standard output.
        # Catch any errors that may occur.
        try:
            canvas.save()
            print("Done")
            print(f"total time: {time.time() - t1}")
        except PermissionError:
            window["output"].update(f"Error: Output file must not be open while saving")
        except OSError:
            window["output"].update(f"Error: Invalid Output Filename")
        # Delete any temporary images that may have been temporarily saved to the current directory through the FadedImage2 class.
        finally:
            for name in my_pic_list:
                if isinstance(name, str):
                    if "MyPica44edtfbdv5yygr" in name:
                        os.remove(name)
            # Once again, print the time it took for the program to run to standard output
            print(f"**{len(my_pic_list)}**")
            print("Done")
            print(f"total time: {time.time()-t1}")
# Close the GUI window
window.close()
