import configparser
import io
import json
import os
import re
import shutil
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import *
from tkinter import filedialog, messagebox, ttk
from tkinter.constants import *

import fitz  # PyMuPDF
from alive_progress import alive_bar
from PIL import Image
from rich import print

import gui
from excel_file import ExcelFile

program_directory = os.path.dirname(os.path.realpath(sys.argv[0]))

global_variables = configparser.ConfigParser()
global_variables.read(f"{program_directory}/global_variables.cfg")

nitrogen_cost_per_hour: int = float(
    global_variables["GLOBAL VARIABLES"]["nitrogen_cost_per_hour"]
)
co2_cost_per_hour: int = float(global_variables["GLOBAL VARIABLES"]["co2_cost_per_hour"])
materials = global_variables["GLOBAL VARIABLES"]["materials"].split(",")
"""
SS      304 SS,409 SS   Nitrogen
ST      Mild Steel      CO2
AL      Aluminium       Nitrogen
"""
gauges = global_variables["GLOBAL VARIABLES"]["gauges"].split(",")
path_to_sheet_prices = global_variables["GLOBAL VARIABLES"]["path_to_sheet_prices"]
with open(path_to_sheet_prices, "r") as f:
    sheet_prices = json.load(f)
size_of_picture = int(global_variables["GLOBAL VARIABLES"]["size_of_picture"])
PROFIT_MARGIN: float = float(global_variables["GLOBAL VARIABLES"]["profit_margin"])
OVERHEAD: float = float(global_variables["GLOBAL VARIABLES"]["overhead"])
path_to_save_quotes = global_variables["GLOBAL VARIABLES"]["path_to_save_quotes"]
path_to_save_workorders = global_variables["GLOBAL VARIABLES"]["path_to_save_workorders"]
price_of_steel_information_path = global_variables["GLOBAL VARIABLES"]["price_of_steel_information"]

with open(price_of_steel_information_path, "r") as f:
    price_of_steel_information = json.load(f)

geofile_name_regex = r"(GEOFILE NAME: [a-zA-z]:\\[\w\W]{1,300}\.geo|GEOFILE NAME: [a-zA-Z]:\\[\w\W]{1,300}\.GEO)"
machining_time_regex = r"(MACHINING TIME: \d{1,}.\d{1,} min)"
weight_regex = r"(WEIGHT: \d{1,}.\d{1,} lb)"
surface_area_regex = r"(SURFACE: \d{1,}.\d{1,}  in2)"
cutting_length_regex = r"(CUTTING LENGTH: \d{1,}.\d{1,}  in|CUTTING LENGTH: \d{1,}  in)"
quantity_regex = r"(  NUMBER: \d{1,})"
part_number_regex = r"(PART NUMBER: \d{1,})"
sheet_quantity_regex = (
    r"(PROGRAMME RUNS:  \/  SCRAP: \d{1,}|PROGRAM RUNS:  \/  SCRAP: \d{1,})"
)
sheet_scrap_percentage_regex = r"(PROGRAMME RUNS:  \/  SCRAP: \d{1,}  \/  \d{1,}.\d{1,} %|PROGRAM RUNS:  \/  SCRAP: \d{1,}  \/  \d{1,}.\d{1,} %)"
scrap_percentage_regex = r"(\d{1,}.\d{1,}) %"
piercing_time_regex = r"(PIERCING TIME \d{1,}.\d{1,}  s)"
material_id_regex = r"MATERIAL ID \(SHEET\): (?=.*(ST|SS|AL)-)"
gauge_regex = r"MATERIAL ID \(SHEET\): .{1,} ?-(\d{1,})"
sheet_dim_regex = r"BLANK: (\d{1,}\.\d{1,} x \d{1,}\.\d{1,}) x \d{1,}\.\d{1,}"

material_selection = ""


class SelectionDialog:
    def __init__(self, parent, choicelist):
        label = ttk.Label(
            parent, text="Select Material Type:"
        )  # .grid(row=0, column=0, sticky="W")
        label.place(rely=1.0, relx=1.0, x=-175, y=-160, anchor=SW, width=150, height=40)

        self.var = StringVar()
        self.var.set("Select Material")  # default option
        self.parent = parent
        popupMenu = ttk.OptionMenu(parent, self.var, *choicelist)
        # popupMenu.grid(sticky=N+S+E+W, row =1, column = 0)
        popupMenu.place(
            rely=1.0, relx=1.0, x=-175, y=-110, anchor=SW, width=150, height=40
        )

        apply_button = ttk.Button(
            parent, text="Apply", command=self.buttonfn
        )  # .grid(row=2, column=0)
        apply_button.place(
            rely=1.0, relx=1.0, x=-175, y=-20, anchor=SW, width=150, height=80
        )

    def buttonfn(self):
        global material_selection
        material_selection = self.var.get()
        self.parent.destroy()


def convert_pdf_to_text(pdf_paths: list, progress_bar) -> None:
    """
    It opens the PDF file, gets the text from each page, and writes it to a text file

    Args:
      pdf_paths (list): list
      progress_bar: a function that will be called after each PDF is processed.
    """
    with open(f"{program_directory}/output.txt", "w") as f:
        f.write("")

    for i, pdf_path in enumerate(pdf_paths, start=1):
        print(f'[ ] Processing "{pdf_path}"\t{i}/{len(pdf_paths)}')
        pdf_file = fitz.open(pdf_path)
        pages = list(range(pdf_file.pageCount))
        for pg in range(pdf_file.pageCount):
            if pg in pages:
                print(f"\t[ ] Getting text from page #{pg+1}.")
                page = pdf_file[pg]
                page_lines = page.get_text("text")
                with open(f"{program_directory}/output.txt", "a") as f:
                    f.write(page_lines)
                print(f"\t[+] Getting text from page #{pg+1}.")
        print(f'[+] Finished "{pdf_path}"\t{i}/{len(pdf_paths)}')
        progress_bar()

    with open(f"{program_directory}/output.txt", "r") as f:
        all_text = f.read()

    with open(f"{program_directory}/output.txt", "w") as f:
        f.write(all_text.replace(" \n", " "))


def extract_images_from_pdf(pdf_paths: list, progress_bar) -> None:
    """
    It opens a PDF file, extracts all the images from it, resizes them to a specific size, and saves
    them to a folder

    Args:
      pdf_paths (list): list = list of paths to the PDF files
      progress_bar: a function that prints a progress bar
    """
    image_count: int = 0
    for i, pdf_path in enumerate(pdf_paths, start=1):
        print(f'[ ] Processing "{pdf_path}"\t{i}/{len(pdf_paths)}')
        pdf_file = fitz.open(pdf_path)
        for page_index in range(len(pdf_file)):
            page = pdf_file[page_index]
            if image_list := page.get_images():
                print(f"\t[+] {len(image_list)} images in page {page_index}")
            else:
                print("\t[!] No images found on page", page_index)
                continue
            print(f"\t[ ] Extracting images from page {page_index}")
            for image_index, img in enumerate(page.get_images(), start=1):
                xref = img[0]
                base_image = pdf_file.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                image = Image.open(io.BytesIO(image_bytes))
                if image.size[0] == 48 and image.size[1] == 48:
                    continue
                image = image.resize(
                    (size_of_picture, size_of_picture), Image.Resampling.LANCZOS
                )
                image.save(
                    open(f"{program_directory}/images/{image_count}.{image_ext}", "wb")
                )
                image_count += 1
            print(f"\t[+] Extracted images from page {page_index}")
        print(f'[+] Finished "{pdf_path}"\t{i}/{len(pdf_paths)}')
        progress_bar()


def convert_material_id_to_name(material: str) -> str:
    """
    It opens the file material_id.json, loads the data, and returns the name of the material

    Args:
      material (str): The material ID of the material you want to convert.

    Returns:
      The name of the material.
    """
    with open(f"{program_directory}/material_id.json", "r") as material_id_file:
        data = json.load(material_id_file)
    return data[material]["name"]


def convert_material_id_to_number(number_id: str) -> str:
    """
    It takes a string as an argument, opens a json file, loads the data from the json file, and returns
    a string

    Args:
      number_id (str): The material ID number.

    Returns:
      The thickness of the material.
    """
    with open(f"{program_directory}/material_id.json", "r") as material_id_file:
        data = json.load(material_id_file)
    return data["thickness"][number_id]


def get_cutting_method(material: str) -> str:
    """
    "Given a material ID, return the cutting method."

    The first line of the function is a docstring. It's a string that describes what the function does.
    It's a good idea to include a docstring in every function you write

    Args:
      material_id (str): The material ID of the material you want to cut.

    Returns:
      The cutting method for the material.
    """
    with open(f"{program_directory}/material_id.json", "r") as material_id_file:
        data = json.load(material_id_file)
    return data[material]["cut"]


def get_table_value_from_text(regex) -> list:
    """
    It takes a regular expression and returns a list of all the matches

    Args:
      regex: The regex to search for.

    Returns:
      A list of all the values in the table.
    """
    with open(f"{program_directory}/output.txt", "r") as f:
        text = f.read()

    items = []

    matches = re.finditer(regex, text, re.MULTILINE)
    for match in matches:
        items.extend(iter(match.groups()))

    return items

def get_value_from_string(regex, text) -> list:
    """
    This function extracts table values from a string using a regular expression and returns them as a
    list.

    Args:
      regex: A regular expression pattern used to match and extract values from the input text.
      text: I'm sorry, but the 'text' parameter is missing. Please provide the text parameter so I can
    assist you better.

    Returns:
      a list of items that match the regular expression pattern in the given text.
    """
    items = []

    matches = re.finditer(regex, text, re.MULTILINE)
    for match in matches:
        items.extend(iter(match.groups()))

    return items

def generate_excel_file(*args, file_name: str):
    """
    It takes in a bunch of lists and generates an excel file with a bunch of data

    Args:
      file_name (str): str = The name of the excel file.
    """
    print("[ ] Generating excel sheet")

    if args[12] == "go":  # Work sheet directory
        excel_document = ExcelFile(
            file_name=f"{path_to_save_workorders}/{file_name}.xlsm",
            program_directory=program_directory
        )
    else:  # Quote directory
        excel_document = ExcelFile(
            file_name=f"{path_to_save_quotes}/{file_name}.xlsm",
            program_directory=program_directory
        )
    # excel_document.create_sheet(sheet_name="info")
    excel_document.add_list_to_sheet(cell="A1", items=materials)
    excel_document.add_list_to_sheet(cell="A2", items=gauges)
    excel_document.add_list_to_sheet(
        cell="A3", items=["Nitrogen", "CO2", "Packing Slip", "Quote", "Work Order"]
    )
    excel_document.add_list_to_sheet(
        cell="A4",
        items=[nitrogen_cost_per_hour, co2_cost_per_hour],
    )
    excel_document.add_list_to_sheet(
        cell="A5",
        items=list(sheet_prices["Price Per Pound"].keys()),
    )
    excel_document.add_list_to_sheet(
        cell="A6",
        items=[
            sheet_prices["Price Per Pound"][sheet_name]["price"]
            for sheet_name in list(sheet_prices["Price Per Pound"].keys())
        ],
    )
    excel_document.set_row_hidden_sheet(cell="A1", hidden=True)
    excel_document.set_row_hidden_sheet(cell="A2", hidden=True)
    excel_document.set_row_hidden_sheet(cell="A3", hidden=True)
    excel_document.set_row_hidden_sheet(cell="A4", hidden=True)
    excel_document.set_row_hidden_sheet(cell="A5", hidden=True)
    excel_document.set_row_hidden_sheet(cell="A6", hidden=True)

    excel_document.add_list_to_sheet(
        cell="A7",
        items=["Total parts: ", "", "", "=ROWS(Table1[Part name])"],
    )
    excel_document.add_list_to_sheet(
        cell="A8",
        items=[
            "Total machine time (min): ",
            "",
            "",
            "=SUMPRODUCT(Table1[Machining time (min)],Table1[Qty])",
            "Total machine time (hour):",
            "",
            "",
            "=$D$6/60",
            "As of: ",
            "=NOW()",
            "done at: ",
            "=NOW()+($D$6/1440)",
        ],
    )
    excel_document.add_list_to_sheet(
        cell="A9",
        items=[
            "Total weight (lb): ",
            "",
            "",
            "=SUMPRODUCT(Table1[Weight (lb)],Table1[Qty])",
        ],
    )
    excel_document.add_list_to_sheet(
        cell="A10",
        items=["Total quantities: ", "", "", "=SUM(Table1[Qty])"],
    )
    excel_document.add_list_to_sheet(
        cell="A11",
        items=[
            "Total surface area (in2): ",
            "",
            "",
            "=SUMPRODUCT(Table1[Surface Area (in2)],Table1[Qty])",
        ],
    )
    excel_document.add_list_to_sheet(
        cell="A12",
        items=[
            "Total cutting length (in): ",
            "",
            "",
            "=SUMPRODUCT(Table1[Cutting Length (in)],Table1[Qty])",
        ],
    )
    excel_document.add_list_to_sheet(
        cell="A13",
        items=[
            "Total piercing time (sec): ",
            "",
            "",
            "=SUMPRODUCT(Table1[Piercing Time (sec)],Table1[Qty])",
        ],
    )
    excel_document.add_item_to_sheet(
        cell="A14",
        item=f"{len(args[5])} files loaded",
    )
    excel_document.add_list_to_sheet(cell="A15", items=args[5], horizontal=False)
    excel_document.add_list_to_sheet(cell=f"A{15+len(args[5])}", items=['Gauge'] + list(price_of_steel_information['pounds_per_square_foot'].keys()), horizontal=True)
    excel_document.add_list_to_sheet(cell=f"A{16+len(args[5])}", items=list(price_of_steel_information['pounds_per_square_foot']['304 SS'].keys()), horizontal=False)
    temp_col = {0: "B", 1: "C",2: "D", 3: "E",4:"F", 5: "G", 6: "H"}
    for i, sheet_name in enumerate(list(price_of_steel_information['pounds_per_square_foot'].keys())):
        for j, thickness in enumerate(price_of_steel_information['pounds_per_square_foot'][sheet_name]):
            excel_document.add_item_to_sheet(cell=f"{temp_col[i]}{16+len(args[5])+j}", item=price_of_steel_information['pounds_per_square_foot'][sheet_name][thickness])

    excel_document.add_image(cell="A1", path_to_image=f"{program_directory}/logo.png")
    excel_document.set_cell_height(cell="A1", height=33)
    excel_document.set_cell_height(cell="A2", height=34)
    excel_document.set_cell_height(cell="A3", height=34)
    if args[12] == 'go':
        excel_document.add_item(cell="E1", item="Work Order")
    else:
        excel_document.add_item(cell="E1", item="Packing Slip")
    excel_document.add_item(cell="E2", item="Order #")
    excel_document.add_list(cell="F1", items=["", "", "", "", "", "", "", "", ""])
    excel_document.add_list(cell="F2", items=["", "", "", "", "", "", "", "", ""])
    excel_document.add_list(cell="F3", items=["", ""])
    excel_document.add_item(cell="A3", item="Date Shipped:")
    excel_document.add_item(cell="E3", item="Ship To:")

    headers = [
        "Item",
        "Part name",
        "Machining time (min)",
        "Weight (lb)",
        "Material",
        "Thickness",
        "Qty",
        "COGS",
        "Overhead",
        "Unit Price",
        "Price",
        "Cutting Length (in)",
        "Surface Area (in2)",
        "Piercing Time (sec)",
        "Total Cost",
    ]

    excel_document.set_cell_width(cell="A1", width=15)
    excel_document.set_cell_width(cell="B1", width=22)
    excel_document.set_cell_width(cell="E1", width=12)
    excel_document.set_cell_width(cell="G1", width=11)
    excel_document.set_cell_width(cell="O1", width=17)
    excel_document.set_cell_width(cell="S1", width=17)
    excel_document.set_cell_width(cell="F1", width=12)
    excel_document.set_cell_width(cell="J1", width=12)
    excel_document.set_cell_width(cell="K1", width=12)
    excel_document.set_cell_width(cell="P1", width=12)
    excel_document.set_cell_width(cell="R1", width=12)

    excel_document.set_col_hidden(cell="C1", hidden=True)
    excel_document.set_col_hidden(cell="D1", hidden=True)
    excel_document.set_col_hidden(cell="H1", hidden=True)
    excel_document.set_col_hidden(cell="I1", hidden=True)
    excel_document.set_col_hidden(cell="L1", hidden=True)
    excel_document.set_col_hidden(cell="M1", hidden=True)
    excel_document.set_col_hidden(cell="N1", hidden=True)
    excel_document.set_col_hidden(cell="O1", hidden=True)

    excel_document.add_item(cell="P2", item="Laser cutting:")
    excel_document.add_item(cell="Q2", item=args[10])
    excel_document.add_dropdown_selection(
        cell="Q2", type="list", location="'info'!$A$3:$B$3"
    )
    excel_document.add_dropdown_selection(
        cell="E1", type="list", location="'info'!$C$3:$E$3"
    )
    STARTING_ROW: int = 5

    for index in range(len(args[0])):
        row: int = index + STARTING_ROW
        excel_document.add_dropdown_selection(
            cell=f"E{row}", type="list", location="'info'!$A$1:$H$1"
        )
        excel_document.add_dropdown_selection(
            cell=f"F{row}", type="list", location="'info'!$A$2:$K$2"
        )

        excel_document.add_item(
            cell=f"B{row}", item=args[0][index]
        )  # File name B
        excel_document.add_item(
            cell=f"C{row}", item=args[1][index]
        )  # Machine Time C
        excel_document.add_item(
            cell=f"D{row}", item=args[2][index]
        )  # Weight D
        excel_document.add_item(
            cell=f"E{row}", item=args[9][index]
        )  # Material Type E
        excel_document.add_item(
            cell=f"F{row}", item=args[8][index]
        )  # Gauge Selection F
        excel_document.add_item(
            cell=f"G{row}", item=args[3][index]
        )  # Quantity G
        excel_document.add_item(
            cell=f"L{row}", item=args[6][index]
        )  # Cutting Length L
        excel_document.add_item(
            cell=f"M{row}", item=args[7][index]
        )  # Surface Area M
        excel_document.add_item(
            cell=f"N{row}", item=args[11][index]
        )  # Piercing Time N
        cost_for_weight = (
            f"INDEX(info!$A$6:$G$6,MATCH($E${row},info!$A$5:$G$5,0))*$D${row}"
        )
        cost_for_time = (
            f"(INDEX('info'!$A$4:$B$4,MATCH($Q$2,'info'!$A$3:$B$3,0))/60)*$C${row}"
        )
        quantity = f"$G{row}"
        excel_document.add_item(
            cell=f"H{row}",
            item=f"=({cost_for_weight}+{cost_for_time})",
            number_format="$#,##0.00",
        )  # Cost

        overhead = f"$J{row}*($T$1)"
        excel_document.add_item(
            cell=f"I{row}",
            item=f"={overhead}",
            number_format="$#,##0.00",
        )  # Overhead

        unit_price = f"CEILING(($O{row})/(1-$T$2),0.01)"
        excel_document.add_item(
            cell=f"J{row}",
            item=f"={unit_price}",
            number_format="$#,##0.00",
        )  # Unit Price

        price = f"CEILING({quantity}*J{row},0.01)"
        excel_document.add_item(
            cell=f"K{row}",
            item=f"={price}",
            number_format="$#,##0.00",
        )  # Price

        total_cost = f"CEILING($H{row}+$I{row},0.01)"
        excel_document.add_item(
            cell=f"O{row}",
            item=f"={total_cost}",
            number_format="$#,##0.00",
        )  # Total Cost

        # Image
        excel_document.add_image(
            cell=f"A{row}",
            path_to_image=f"{program_directory}/images/{args[4][index]}.jpeg",
        )

        excel_document.set_cell_height(cell=f"A{row}", height=78)

    excel_document.add_table(
        display_name="Table1",
        theme="TableStyleLight8",
        location=f"A{STARTING_ROW-1}:O{index+STARTING_ROW}",
        headers=headers,
    )
    excel_document.add_item(cell=f"A{index+STARTING_ROW+1}", item="", totals=True)
    excel_document.add_item(cell=f"B{index+STARTING_ROW+1}", item="", totals=True)
    excel_document.add_item(cell=f"C{index+STARTING_ROW+1}", item="", totals=True)
    excel_document.add_item(cell=f"D{index+STARTING_ROW+1}", item="", totals=True)
    excel_document.add_item(cell=f"E{index+STARTING_ROW+1}", item="", totals=True)
    excel_document.add_item(cell=f"F{index+STARTING_ROW+1}", item="", totals=True)
    excel_document.add_item(cell=f"P{index+STARTING_ROW}", item=f"Scrap: {args[14]}%", totals=False)
    excel_document.add_item(cell=f"P{index+STARTING_ROW+1}", item=f"Sheets:", totals=False)
    excel_document.add_item(cell=f"Q{index+STARTING_ROW+1}", item=args[13], totals=False)
    excel_document.add_item(
        cell=f"C{index+STARTING_ROW+1}",
        item="=SUMPRODUCT(Table1[Machining time (min)],Table1[Qty])",
        totals=True,
    )
    excel_document.add_item(
        cell=f"D{index+STARTING_ROW+1}",
        item="=SUMPRODUCT(Table1[Weight (lb)],Table1[Qty])",
        totals=True,
    )
    excel_document.add_item(cell=f"R{index+STARTING_ROW+1}", item="Total:", totals=False)
    sheet_dim_left = f'TEXTAFTER("{args[15]}", "x")'
    sheet_dim_right = f'TEXTBEFORE("{args[15]}", "x")'
    price_per_pound = "INDEX(info!$A$6:$G$6,MATCH($E$6, info!$A$5:$G$5,0))"
    pounds_per_sheet = f'INDEX(info!$B${16+len(args[5])}:$H${16+len(args[5])+15},MATCH($F$6,info!$A${16+len(args[5])}:$A${16+len(args[5])+15},0),MATCH($E$6,info!$B${15+len(args[5])}:$H${15+len(args[5])},0))'
    sheet_quantity = f'Q{index+STARTING_ROW+1}'
    excel_document.add_item(cell=f"S{index+STARTING_ROW+1}", item=f'={sheet_dim_right}*{sheet_dim_left}/144*{price_per_pound}*{pounds_per_sheet}*{sheet_quantity}', number_format="$#,##0.00", totals=False)
    excel_document.add_item(cell=f"G{index+STARTING_ROW+1}", item="", totals=True)
    excel_document.add_item(cell=f"J{index+STARTING_ROW+1}", item="Total: ", totals=True)
    excel_document.add_item(
        cell=f"H{index+STARTING_ROW+1}",
        item="=SUM(Table1[COGS])",
        number_format="$#,##0.00",
        totals=True,
    )
    excel_document.add_item(
        cell=f"I{index+STARTING_ROW+1}",
        item="=SUM(Table1[Overhead])",
        number_format="$#,##0.00",
        totals=True,
    )
    excel_document.add_item(
        cell=f"K{index+STARTING_ROW+1}",
        item="=SUM(Table1[Price])",
        number_format="$#,##0.00",
        totals=True,
    )
    excel_document.add_item(cell=f"K{index+STARTING_ROW+2}", item="No Tax Included")
    excel_document.add_item(
        cell=f"B{index+STARTING_ROW+2}",
        item="Payment past due date will receive 1.5% interest rate per month of received goods.",
    )
    excel_document.add_item(
        cell=f"L{index+STARTING_ROW+1}",
        item="=SUMPRODUCT(Table1[Cutting Length (in)],Table1[Qty])",
        totals=True,
    )
    excel_document.add_item(
        cell=f"M{index+STARTING_ROW+1}",
        item="=SUMPRODUCT(Table1[Surface Area (in2)],Table1[Qty])",
        totals=True,
    )
    excel_document.add_item(
        cell=f"N{index+STARTING_ROW+1}",
        item="=SUMPRODUCT(Table1[Piercing Time (sec)],Table1[Qty])",
        totals=True,
    )
    excel_document.add_item(
        cell=f"O{index+STARTING_ROW+1}",
        item="=SUM(Table1[Total Cost])",
        totals=True,
    )
    excel_document.add_item(
        cell=f"A{index+STARTING_ROW+4}",
        item="Date expected:",
    )
    excel_document.add_item(
        cell=f"A{index+STARTING_ROW+6}",
        item="_______________________",
    )
    excel_document.add_item(
        cell=f"E{index+STARTING_ROW+4}",
        item="Received in good order by:",
    )
    excel_document.add_item(
        cell=f"E{index+STARTING_ROW+6}",
        item="______________________________",
    )

    excel_document.add_item(cell="S1", item="Overhead:")
    excel_document.add_item(cell="T1", item=OVERHEAD, number_format="0%")

    excel_document.add_item(cell="S2", item="Profit Margin:")
    excel_document.add_item(cell="T2", item=PROFIT_MARGIN, number_format="0%")

    excel_document.set_print_area(cell=f"A1:K{index + STARTING_ROW+6}")

    print("\t[ ] Injecting macro.bin")
    excel_document.add_macro(macro_path=f"{program_directory}/macro.bin")

    if args[12] == "go":
        excel_document.set_col_hidden("J1", True)
        excel_document.set_col_hidden("K1", True)

    print("\t[+] Injected macro.bin")
    excel_document.save()
    print("[+] Excel sheet generated.")
    os.remove("action")


def save_json_file(dictionary: dict, file_name: str) -> None:
    """
    It takes a dictionary and a file name as arguments, and then saves the dictionary as a json file
    with the given file name

    Args:
      dictionary (dict): The dictionary you want to save.
      file_name (str): The name of the file you want to save.
    """
    with open(f"{program_directory}/excel files/{file_name}.json", "w") as fp:
        json.dump(dictionary, fp, sort_keys=True, indent=4)


def convert(file_names: list):  # sourcery skip: low-code-quality
    """
    It takes a list of file names, extracts the images from the PDFs, converts the PDFs to text,
    extracts the data from the text, and then generates an excel file and a JSON file

    Args:
      file_names (list): list
    """

    choicewin = tk.Tk()
    choicewin.resizable(False, False)
    choicewin.lift()
    choicewin.attributes("-topmost", True)
    width, height = 200, 200
    choicewin.geometry(f"{width}x{height}")
    choicewin.minsize(width, height)
    choicewin.maxsize(width, height)
    choicewin.title("Choose Material")
    t_materials = materials
    t_materials.insert(0, "304 SS")
    app = SelectionDialog(choicewin, t_materials)
    choicewin.mainloop()
    if material_selection == "":
        return

    Path(f"{program_directory}/images").mkdir(parents=True, exist_ok=True)
    today = datetime.now()
    current_time = today.strftime("%Y-%m-%d-%H-%M-%S")

    with alive_bar(
        2 + (len(file_names) * 4),
        dual_line=True,
        title="Generating",
        force_tty=True,
        theme="smooth",
    ) as progress_bar:
        part_dictionary = {}
        part_names = []
        quantity_numbers = []
        machining_times_numbers = []
        weights_numbers = []
        part_numbers = []
        surface_areas_numbers = []
        cutting_lengths_numbers = []
        piercing_time_numbers = []
        material_for_parts = []
        gauge_for_parts = []
        quantity_multiplier: int = 1
        scrap_percentage: float = 0
        total_sheet_count: int = 0

        progress_bar.text = "-> Getting images, please wait..."
        extract_images_from_pdf(file_names, progress_bar)
        progress_bar()

        for file_name in file_names:
            progress_bar.text = "-> Getting text, please wait..."

            progress_bar.text = "-> Getting all data, please wait..."
            convert_pdf_to_text([file_name], progress_bar)
            progress_bar()

            quantity_multiplier = get_table_value_from_text(regex=sheet_quantity_regex)[0]
            quantity_multiplier = int(
                quantity_multiplier.replace("PROGRAMME RUNS:  /  SCRAP: ", "").replace(
                    "PROGRAM RUNS:  /  SCRAP: ", ""
                )
            )
            total_sheet_count += quantity_multiplier
            scrap_percentage_string = get_table_value_from_text(regex=sheet_scrap_percentage_regex)[0]
            scrap_percentage: float = float(get_value_from_string(regex=scrap_percentage_regex, text=scrap_percentage_string)[0])
            sheet_dim: str = get_table_value_from_text(regex=sheet_dim_regex)[0].replace(
                " ", ""
            )
            # material_for_part = convert_material_id_to_name(
            # material=get_table_value_from_text(regex=material_id_regex)[0]
            # )
            material_for_part = material_selection
            gauge_for_part = convert_material_id_to_number(
                number_id=get_table_value_from_text(regex=gauge_regex)[0],
            )

            if (
                int(get_table_value_from_text(regex=gauge_regex)[0]) >= 50
            ):  # More than 1/2 inch
                material_for_part = "Laser Grade Plate"
            cutting_with = get_cutting_method(
                material=get_table_value_from_text(regex=material_id_regex)[0]
            )
            part_dictionary["_" + file_name] = {
                "quantity_multiplier": quantity_multiplier,
                "gauge": gauge_for_part,
                "material": material_for_part,
                "sheet_dim": sheet_dim,
            }
            part_file_paths = get_table_value_from_text(regex=geofile_name_regex)

            for part_name in part_file_paths:
                part_name = (
                    part_name.split("\\")[-1]
                    .replace("\n", "")
                    .replace(".GEO", "")
                    .strip()
                )
                part_dictionary[part_name] = {
                    "quantity": 0,
                    "machine_time": 0.0,
                    "weight": 0.0,
                    "part_number": 0,
                    "image_index": 0,
                    "surface_area": 0,
                    "cutting_length": 0,
                    "file_name": file_name,
                    "piercing_time": 0.0,
                    "gauge": gauge_for_part,
                    "material": material_for_part,
                    "recut": False,
                    "sheet_dim": sheet_dim,
                }

                part_names.append(part_name)

            quantities = get_table_value_from_text(regex=quantity_regex)
            for quantity in quantities:
                quantity = quantity.replace("  NUMBER: ", "")
                quantity_numbers.append((int(quantity) * quantity_multiplier))

            machining_times = get_table_value_from_text(regex=machining_time_regex)
            for machining_time in machining_times:
                machining_time = machining_time.replace("MACHINING TIME: ", "").replace(
                    " min", ""
                )
                machining_times_numbers.append(float(machining_time))

            weights = get_table_value_from_text(regex=weight_regex)
            for weight in weights:
                weight = weight.replace("WEIGHT: ", "").replace(" lb", "")
                weights_numbers.append(float(weight))

            surface_areas = get_table_value_from_text(regex=surface_area_regex)
            for surface_area in surface_areas:
                surface_area = surface_area.replace("SURFACE: ", "").replace("  in2", "")
                surface_areas_numbers.append(float(surface_area))

            cutting_lengths = get_table_value_from_text(regex=cutting_length_regex)
            for cutting_length in cutting_lengths:
                cutting_length = cutting_length.replace("CUTTING LENGTH: ", "").replace(
                    "  in", ""
                )
                cutting_lengths_numbers.append(float(cutting_length))

            piercing_times = get_table_value_from_text(regex=piercing_time_regex)
            for piercing_time in piercing_times:
                piercing_time = piercing_time.replace("PIERCING TIME ", "").replace(
                    "  s", ""
                )
                piercing_time_numbers.append(float(piercing_time))

            part_numbers_string = get_table_value_from_text(regex=part_number_regex)
            for part_number in part_numbers_string:
                part_number = part_number.replace("PART NUMBER: ", "")
                part_numbers.append(int(part_number))

        for i, part_name in enumerate(part_names):
            part_dictionary[part_name]["quantity"] += quantity_numbers[i]
            part_dictionary[part_name]["machine_time"] = machining_times_numbers[i]
            part_dictionary[part_name]["weight"] = weights_numbers[i]
            part_dictionary[part_name]["part_number"] = part_numbers[i]
            part_dictionary[part_name]["image_index"] = i
            part_dictionary[part_name]["surface_area"] = surface_areas_numbers[i]
            part_dictionary[part_name]["cutting_length"] = cutting_lengths_numbers[i]
            part_dictionary[part_name]["piercing_time"] = piercing_time_numbers[i]

        part_names.clear()
        machining_times_numbers.clear()
        part_numbers.clear()
        quantity_numbers.clear()
        weights_numbers.clear()
        piercing_time_numbers.clear()
        surface_areas_numbers.clear()
        cutting_lengths_numbers.clear()
        image_index = []

        save_json_file(dictionary=part_dictionary, file_name=current_time)

        gui.load_gui(
            f"{program_directory}/excel files/{current_time}.json", material_selection
        )
        with open(f"{program_directory}/excel files/{current_time}.json", "r") as f:
            part_dictionary = json.load(f)

        for part_name in part_dictionary:
            if part_name[0] != "_":
                part_names.append(part_name)

        for part in part_dictionary:
            try:
                machining_times_numbers.append(part_dictionary[part]["machine_time"])
                part_numbers.append(part_dictionary[part]["part_number"])
                quantity_numbers.append(part_dictionary[part]["quantity"])
                weights_numbers.append(part_dictionary[part]["weight"])
                image_index.append(part_dictionary[part]["image_index"])
                material_for_parts.append(part_dictionary[part]["material"])
                gauge_for_parts.append(part_dictionary[part]["gauge"])
                piercing_time_numbers.append(part_dictionary[part]["piercing_time"])
                cutting_lengths_numbers.append(part_dictionary[part]['cutting_length'])
                surface_areas_numbers.append(part_dictionary[part]['surface_area'])
            except KeyError:
                continue

        progress_bar.text = "-> Generating excel sheet, please wait..."
        progress_bar()

        try:
            with open(f"{program_directory}/action", "r") as f:
                action = f.read()
        except Exception:
            return

        generate_excel_file(
            part_names,                 #0
            machining_times_numbers,    #1
            weights_numbers,            #2
            quantity_numbers,           #3
            image_index,                #4
            file_names,                 #5
            surface_areas_numbers,      #6
            cutting_lengths_numbers,    #7
            gauge_for_parts,            #8
            material_for_parts,         #9
            cutting_with,               #10
            piercing_time_numbers,      #11
            action,                     #12 'go' or 'quote'
            total_sheet_count,          #13
            scrap_percentage,           #14
            sheet_dim,                  #15
            file_name=current_time,
        )

        print(f'Opening "{program_directory}/excel files/{current_time}.xlsm"')

        progress_bar()
        progress_bar.text = "-> Finished! :)"

        if action == 'go':
            os.startfile(f'"{path_to_save_workorders}/{current_time}.xlsm"')
        shutil.rmtree(f"{program_directory}/images")


file_names: str = sys.argv[-1].split("\\")[-1]
directory_of_file: str = os.getcwd()

root = tk.Tk()
root.withdraw()


Path(f"{program_directory}/excel files").mkdir(parents=True, exist_ok=True)


filetypes = (("pdf files", "*.pdf"),)
file_paths = filedialog.askopenfilenames(
    parent=root, title="Select files", initialdir=directory_of_file, filetypes=filetypes
)

if len(file_paths) > 0:
    root.destroy()
    convert(file_paths)
