import pyautogui as pt
from PIL import Image
import pytesseract
import xlsxwriter
import openpyxl
import os
import regex
import shutil
import logging

class Overcuts:
    def __int__(self):
        # Locations of the Required Files
        self.LOCATION = os.path.abspath(os.path.dirname(__file__))  # Project Folder
        self.excel_file = self.LOCATION + r"\data\overcuts.xlsx"  # Excel File that will store the Parsed Data
        self.excel_file_v2_original = self.LOCATION + r"\data\overcuts-v2.xlsx"  # V2 Original Excel File
        self.excel_file_v2 = self.LOCATION + r"\data\overcuts-temp.xlsx"  # V2 temp excel file
        self.img_distance_field = self.LOCATION + r"\data\distance_field_laptop.png"  # Distance Field image to location
        self.pytesseract = pytesseract
        self.pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\Tesseract.exe"  # Tesseract OCR
        self.img1 = "data/dd1.png"  # Template image - 1 to store new created image from screenshot
        self.img2 = "data/dd2.png"  # Template image - 2
        self.img3 = "data/dd3.png"  # Template image - 3
        self.img4 = "data/dd4.png"  # Template image - 4
        self.row = 0  # Starting Row Number for Excel
        self.col = 0  # Starting Column Number for Excel

    # Function to start an Excel file which will save all the required overcuts data
    # Failing to start this function will result in error
    def start_excel(self):
        try:
            self.workbook = xlsxwriter.Workbook(self.excel_file)
            self.worksheet = self.workbook.add_worksheet()
        except Exception as e:
            print(f"""Error : {type(e).__name__}
            Location: Starting Excel File
            File Location: {self.excel_file}""")
            exit()
        # Entering Data into Header Row
        # Text Formatting for Header Row
        head_style = {'bold': True, 'font_color': 'grey'}
        header_fields = ("Distance", "Surface", "Elevation", "Depth", "Calculated Depth", "Tile")
        for index, value in enumerate(header_fields):
            self.add_to_excel(0, index, value, head_style)
        self.row += 1

    # Starts V2 Excel File
    def start_excel_v2(self):
        os.remove(self.excel_file_v2)  # Remove Exisiting V2 Temp File
        shutil.copy(self.excel_file_v2_original, self.excel_file_v2)  # Create New V2 Temp File

        self.oxl = openpyxl.load_workbook(self.excel_file_v2)
        self.xl_sheet = self.oxl['M1']
        self.row = 6

    # This function adds values to Excel files at provided Row and Column Location
    def add_to_excel(self, row_value, col_value, input_value, input_format):
        # Text Formatting for Input
        data_format = self.workbook.add_format(input_format)
        # Filling Data
        try:
            self.worksheet.write(row_value, col_value, input_value, data_format)
        except Exception as e:
            print(f"""Error : {type(e).__name__}
                        Location: Entering Data into Excel File""")
            exit()

    # V2 - adds values to Excel files at provided Row and Column Location
    def add_to_excel_v2(self, input_cell, input_value):
        try:
            self.xl_sheet[input_cell] = input_value
        except Exception as e:
            print(f"""Error : {type(e).__name__}
                        Location: V2 - Entering Data into Excel File""")
            exit()

    # Close Excel File
    def close_excel(self):
        self.workbook.close()

    # v2 - Close Excel File
    def close_excel_v2(self):
        # Formula to Remove Empty Rows. Did not work as expected
        ''' self.xl_sheet.delete_rows(self.row, 40 - self.row)
        for row in self.xl_sheet.iter_rows(min_row=self.row, max_row=40):
            for cell in row:
                cell.value = None '''
        self.oxl.save(self.excel_file_v2)

    # This function will open the Excel file
    def show_result(self):
        os.startfile(self.excel_file)

    # v2 - This function will open the Excel file
    def show_result_v2(self):
        os.startfile(self.excel_file_v2)

    # This function is used to find the location of an image on screen
    # Arguments: image_location=Location of image that to be located,
    # probability=Similarity of image to be located to the image found on desktop, Value=0 to 1
    @staticmethod
    def locate_image(image_location, probability=0.7):
        try:
            position = pt.locateOnScreen(image_location, confidence=probability)
            x = position[0]
            y = position[1]
            return x, y
        except Exception as e:  # if image not found
            print(f"""Error : {type(e).__name__}
            Location: Locating Image {image_location}""")
            exit()

    # This function locates the image of distance field using locate_image function
    def locate_distance_field(self):
        coordinates = self.locate_image(self.img_distance_field, 0.7)
        self.x_coordinate = coordinates[0]
        self.y_coordinate = coordinates[1]

        # Try and except case in case mouse cursor is unable to move to the coordinates location
        try:
            pt.moveTo(self.x_coordinate, self.y_coordinate, duration=.05)
        except Exception as e:
            print(f"""Error : {type(e).__name__}
            Location: Moving Mouse Cursor to Distance Field""")
            exit()

    # This function moves the cursor to Entered distance value
    # which helps to read the required values from screen
    def move_to_distance(self, input_distance):
        pt.doubleClick(self.x_coordinate + 110, self.y_coordinate + 12, interval=0)
        pt.press('backspace')

        # Converting input_distance integer to string to enter each character one by one into the input field
        text_to_str = str(input_distance)
        k = 0
        while k < len(str(text_to_str)):
            pt.press(text_to_str[k])
            k += 1

    # Take a snip out of screen and converts that snip of image into text using Tesseract OCR
    def snip_to_text(self, configuration, add_to_x=0, add_to_y=0, image_size_x=67, image_size_y=17, output_type="text"):
        try:
            im = pt.screenshot(region=(self.x_coordinate + add_to_x, self.y_coordinate - add_to_y, image_size_x, image_size_y))
            im.save(self.img1)
            text = pytesseract.image_to_string(Image.open(self.img1), lang="eng", config=configuration)
            print("OCR: " + text)
        except Exception as e:
            print(f"""Error : {type(e).__name__}
            Location: OCR""")
            exit()

        if output_type == "number":
            # Cleaning Text - Removing White Spaces and Non-Alphanumeric Symbols
            # Test - print("Cleaning Number: " + text + "\n")
            text = text.replace('O', '0').replace('o', '0').replace('\n', '')
            text = float(regex.sub(r"([^\d.]|(?<=\..*)\.)", "", text))
        return text


# User Inputs (Start and End) to start Overcut Script
starting_distance = int(input('Enter Starting Distance for Overcut: '))
stop_distance = int(input('Enter End Distance for Overcut: '))

# Starting Overcut
overcut_start = Overcuts()
overcut_start.__int__()
#overcut_start.start_excel()
overcut_start.start_excel_v2()
overcut_start.locate_distance_field()
number_format = {'num_format': '#,##0.00'}
while starting_distance <= stop_distance:
    overcut_start.move_to_distance(starting_distance)

    # V1 SYSTEM HERE
    """ overcut_start.add_to_excel(overcut_start.row, 0, starting_distance, number_format)  # Starting Distance to Excel
    surface_val = overcut_start.snip_to_text('--psm 7', 66, 95, 67, 17, "number")
    overcut_start.add_to_excel(overcut_start.row, 1, surface_val, number_format)  # Surface Value to Excel
    elevation_val = overcut_start.snip_to_text('--psm 7', 66, 71, 67, 17, "number")
    overcut_start.add_to_excel(overcut_start.row, 2, elevation_val, number_format)  # Elevation Value to Excel
    depth_val = overcut_start.snip_to_text('--psm 7', 66, 49, 53, 19, "number")
    overcut_start.add_to_excel(overcut_start.row, 3, depth_val, number_format)  # Depth Value to Excel
    tile_size_val = overcut_start.snip_to_text('--psm 6', 66, 119, 56, 17, "text")
    overcut_start.add_to_excel(overcut_start.row, 5, tile_size_val, '')  # Tile Size Value to Excel 
    
    # Depth Calculation to verify it with captured depth
    calculated_depth = round(float(surface_val - elevation_val), 2)
    if calculated_depth != depth_val:
        bold_format = {'bold': True, 'font_color': 'red'}
        overcut_start.add_to_excel(overcut_start.row, 4, calculated_depth, bold_format)
    else:
        overcut_start.add_to_excel(overcut_start.row, 4, calculated_depth, number_format)
    """

    # V2 SYSTEM HERE
    overcut_start.add_to_excel_v2("C" + str(overcut_start.row), starting_distance)  # Starting Distance to Excel
    surface_val = overcut_start.snip_to_text('--psm 7', 66, 95, 67, 17, "number")
    overcut_start.add_to_excel_v2("F" + str(overcut_start.row), surface_val)  # Surface Value to Excel
    elevation_val = overcut_start.snip_to_text('--psm 7', 66, 71, 67, 17, "number")
    overcut_start.add_to_excel_v2("G" + str(overcut_start.row), elevation_val)  # Elevation Value to Excel
    depth_val = overcut_start.snip_to_text('--psm 7', 66, 49, 53, 19, "number")
    overcut_start.add_to_excel_v2("N" + str(overcut_start.row), depth_val)  # Depth Value to Excel
    tile_size_val = overcut_start.snip_to_text('--psm 6', 66, 119, 56, 17, "number")
    overcut_start.add_to_excel_v2("C" + str(overcut_start.row), tile_size_val)  # Tile Size Value to Excel

    # Depth Calculation to verify it with captured depth
    calculated_depth = round(float(surface_val - elevation_val), 2)
    if calculated_depth != depth_val:
        bold_format = {'bold': True, 'font_color': 'red'}
        overcut_start.add_to_excel_v2("O" + str(overcut_start.row), calculated_depth)
    else:
        overcut_start.add_to_excel_v2("O" + str(overcut_start.row), calculated_depth)

    overcut_start.row += 1
    starting_distance = starting_distance + 100

#overcut_start.close_excel()
#overcut_start.show_result()

overcut_start.close_excel_v2()
overcut_start.show_result_v2()
