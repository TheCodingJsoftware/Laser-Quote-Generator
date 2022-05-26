import configparser
import contextlib
import io
import os
import re
import shutil
import sys
import tkinter as tk
from pathlib import Path
from time import sleep
from tkinter import filedialog

import fitz  # PyMuPDF
import openpyxl
from alive_progress import alive_bar
from openpyxl import Workbook
from openpyxl.drawing import image
from openpyxl.styles import Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from PIL import Image
from rich import print

program_directory = os.path.dirname(os.path.realpath(sys.argv[0]))

global_variables = configparser.ConfigParser()
global_variables.read(f"{program_directory}/global_variables.cfg")

nitrogen_cost_per_hour: int = float(
    global_variables["GLOBAL VARIABLES"]["nitrogen_cost_per_hour"]
)
co2_cost_per_hour: int = float(global_variables["GLOBAL VARIABLES"]["co2_cost_per_hour"])
materials = global_variables["GLOBAL VARIABLES"]["materials"].split(",")
gauges = global_variables["GLOBAL VARIABLES"]["gauges"].split(",")
path_to_sheet_prices = global_variables["GLOBAL VARIABLES"]["path_to_sheet_prices"]


geofile_name_regex = r"(GEOFILE NAME: C:\\[\w\W]{1,300}\.GEO)"
machining_time_regex = r"(MACHINING TIME: \d{1,}.\d{1,} min)"
weight_regex = r"(WEIGHT: \d{1,}.\d{1,} lb)"
quantity_regex = r"(  NUMBER: \d{1,})"
part_number_regex = r"(PART NUMBER: \d{1,})"


def convert_pdf_to_text(pdf_paths: list, bar) -> None:
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
        bar()

    with open(f"{program_directory}/output.txt", "r") as f:
        all_text = f.read()

    with open(f"{program_directory}/output.txt", "w") as f:
        f.write(all_text.replace(" \n", " "))


def extract_images_from_pdf(pdf_paths: list, bar) -> None:
    image_count: int = 0
    for i, pdf_path in enumerate(pdf_paths, start=1):
        print(f'[ ] Processing "{pdf_path}"\t{i}/{len(pdf_paths)}')
        pdf_file = fitz.open(pdf_path)
        for page_index in range(len(pdf_file)):
            page = pdf_file[page_index]
            if image_list := page.getImageList():
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
                image = image.resize((75, 75), Image.ANTIALIAS)
                image.save(
                    open(f"{program_directory}/images/{image_count}.{image_ext}", "wb")
                )
                image_count += 1
            print(f"\t[+] Extracted images from page {page_index}")
        print(f'[+] Finished "{pdf_path}"\t{i}/{len(pdf_paths)}')
        bar()


def get_table_value_from_text(regex) -> list:
    with open(f"{program_directory}/output.txt", "r") as f:
        text = f.read()

    items = []

    matches = re.finditer(regex, text, re.MULTILINE)
    for match in matches:
        items.extend(iter(match.groups()))
    return items


def generate_excel_file(*args):
    print("[ ] Generating excel sheet")
    wb = Workbook()
    wb.create_sheet("Sheet 2")
    ws = wb.active
    source = wb.get_sheet_by_name("Sheet 2")
    source.append(materials)
    source.append(gauges)
    source.append(["Nitrogen", "CO2"])
    source.append([nitrogen_cost_per_hour, co2_cost_per_hour])
    num: int = 0
    headers = [
        "",
        "File name:",
        "Part #:",
        "Machining time (min):",
        "Weight (lb):",
        "Quantity",
        "Material Type",
        "Gauge",
        "Cost",
        "Laser cutting:",
        "Nitrogen",
    ]
    ws.append(headers)

    ws.column_dimensions["A"].width = 11
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 10
    ws.column_dimensions["G"].width = 20
    ws.column_dimensions["j"].width = 14

    material_selection = DataValidation(type="list", formula1="'Sheet 2'!$A$1:$H$1")
    ws.add_data_validation(material_selection)

    gauge_selection = DataValidation(type="list", formula1="'Sheet 2'!$A$2:$K$2")
    ws.add_data_validation(gauge_selection)

    laser_cutting = DataValidation(type="list", formula1="'Sheet 2'!$A$3:$B$3")
    ws.add_data_validation(laser_cutting)
    laser_cutting.add(ws["$K$1"])

    for part_number, machine_time, weight in zip(args[0], args[1], args[2]):
        row: int = num + 2
        ws.append(
            [
                "",
                args[0][num],
                args[1][num],
                args[2][num],
                args[3][num],
                args[4][num],
                materials[0],
                gauges[0],
            ]
        )
        material_selection.add(ws[f"G{row}"])

        gauge_selection.add(ws[f"H{row}"])

        cost_for_weight = f"INDEX('{path_to_sheet_prices}'!$D$6:$J$6,MATCH(G{row},'{path_to_sheet_prices}'!$D$5:$J$5,0))*$E{row}"
        cost_for_time = f"(INDEX('Sheet 2'!A4:B4,MATCH(K1,'Sheet 2'!A3:B3,0))/60)*$D{row}"
        quantity = f"$F{row}"
        ws[f"I{ row}"] = f"=({cost_for_weight}+{cost_for_time})*{quantity}"

        for col in [2, 3, 4, 5, 6, 7, 8, 9]:
            ws.cell(row, col).alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )

        _cell = ws.cell(row, 9)
        _cell.number_format = "$#,##0.00"

        img = image.Image(f"{program_directory}/images/{num}.jpeg")
        img.anchor = f"A{row}"
        ws.row_dimensions[row].height = 57
        ws.add_image(img)
        num += 1

    tab = Table(displayName="Table1", ref=f"B1:I{num+1}")

    style = TableStyleInfo(
        name="TableStyleLight9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=True,
    )
    tab.tableStyleInfo = style
    ws.add_table(tab)

    for i, j in enumerate(args):
        if i > 4:
            ws.append(args[i])

    _cell = ws.cell(num + 6, 3)
    _cell.number_format = "$#,##0.00"

    wb.save(f"{program_directory}/excel_sheet.xlsx")  # save to excel file.
    print("[+] Excel sheet generated.")


def convert(file_names: list):
    try:
        if Path(f"{program_directory}/excel_sheet.xlsx").is_file():
            os.remove(f"{program_directory}/excel_sheet.xlsx")
    except Exception:
        print("You have this excel spread sheet open, close it and try again.")
        sleep(5)
        return

    Path(f"{program_directory}/images").mkdir(parents=True, exist_ok=True)

    with alive_bar(
        4 + (len(file_names) * 2),
        dual_line=True,
        title="Generating",
        force_tty=True,
        theme="smooth",
    ) as bar:
        bar.text = "-> Getting text, please wait..."

        convert_pdf_to_text(file_names, bar)

        bar.text = "-> Getting images, please wait..."
        bar()

        extract_images_from_pdf(file_names, bar)

        bar.text = "-> Generating excel sheet, please wait..."
        bar()

        part_file_paths = get_table_value_from_text(regex=geofile_name_regex)
        file_names = [
            part_file_path.split("\\")[-1].replace("\n", "").replace(".GEO", "")
            for part_file_path in part_file_paths
        ]

        quantity = get_table_value_from_text(regex=quantity_regex)
        quantity_numbers = [int(time.replace("  NUMBER: ", "")) for time in quantity]

        machining_times = get_table_value_from_text(regex=machining_time_regex)
        machining_times_numbers = [
            float(time.replace("MACHINING TIME: ", "").replace(" min", ""))
            for time in machining_times
        ]

        weights = get_table_value_from_text(regex=weight_regex)
        weights_numbers = [
            float(time.replace("WEIGHT: ", "").replace(" lb", "")) for time in weights
        ]

        part_numbers_string = get_table_value_from_text(regex=part_number_regex)
        part_numbers = [
            int(time.replace("PART NUMBER: ", "")) for time in part_numbers_string
        ]

        generate_excel_file(
            file_names,
            part_numbers,
            machining_times_numbers,
            weights_numbers,
            quantity_numbers,
            [],
            ["", "Total time (min):", "=SUM(Table1[Machining time (min):])"],
            ["", "Total weight (lb):", "=SUM(Table1[Weight (lb):])"],
            ["", "Total parts:", "=SUM(Table1[Quantity])"],
            ["", "Total cost:", "=SUM(Table1[Cost])"],
        )
        bar()

        print(f'Opening "{program_directory}/excel_sheet.xlsx"')

        bar()

        os.startfile(f'"{program_directory}/excel_sheet.xlsx"')

        shutil.rmtree(f"{program_directory}/images")


file_names: str = sys.argv[-1].split("\\")[-1]
directory_of_file: str = os.getcwd()

root = tk.Tk()
root.withdraw()

filetypes = (("pdf files", "*.pdf"),)

file_paths = filedialog.askopenfilenames(
    parent=root, title="Select files", initialdir=directory_of_file, filetypes=filetypes
)

if len(file_paths) > 0:
    root.destroy()
    convert(file_paths)
