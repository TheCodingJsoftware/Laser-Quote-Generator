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

from excel_file import ExcelFile

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
size_of_picture = int(global_variables["GLOBAL VARIABLES"]["size_of_picture"])


geofile_name_regex = r"(GEOFILE NAME: C:\\[\w\W]{1,300}\.GEO)"
machining_time_regex = r"(MACHINING TIME: \d{1,}.\d{1,} min)"
weight_regex = r"(WEIGHT: \d{1,}.\d{1,} lb)"
quantity_regex = r"(  NUMBER: \d{1,})"
part_number_regex = r"(PART NUMBER: \d{1,})"
sheet_quantity_regex = r"(PROGRAMME RUNS:  \/  SCRAP: \d{1,})"


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
                image = image.resize((size_of_picture, size_of_picture), Image.ANTIALIAS)
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

    excel_document = ExcelFile(file_name=f"{program_directory}/excel_sheet.xlsx")

    excel_document.create_sheet(sheet_name="Sheet 2")

    excel_document.add_list_to_sheet(sheet_name="Sheet 2", cell="A1", items=materials)
    excel_document.add_list_to_sheet(sheet_name="Sheet 2", cell="A2", items=gauges)
    excel_document.add_list_to_sheet(
        sheet_name="Sheet 2", cell="A3", items=["Nitrogen", "CO2"]
    )
    excel_document.add_list_to_sheet(
        sheet_name="Sheet 2",
        cell="A4",
        items=[nitrogen_cost_per_hour, co2_cost_per_hour],
    )

    excel_document.add_image(
        cell="A1", path_to_image=f"{program_directory}/Piney MGF Logo.png"
    )
    excel_document.set_cell_height(cell="A1", height=67)

    headers = [
        "Part name",
        "Machining time (min)",
        "Weight (lb)",
        "Quantity",
        "Material",
        "Gauge",
        "Price ($)",
    ]

    excel_document.add_item(cell="G1", item="Quote Name:")
    excel_document.add_list(cell="B2", items=headers)

    excel_document.set_cell_width(cell="A1", width=size_of_picture / 6)
    excel_document.set_cell_width(cell="B1", width=22)
    excel_document.set_col_hidden(cell="C1", hidden=True)
    excel_document.set_col_hidden(cell="D1", hidden=True)
    excel_document.set_cell_width(cell="O1", width=15)
    excel_document.set_cell_width(cell="G1", width=15)
    excel_document.set_cell_width(cell="H1", width=15)

    excel_document.add_item(cell="O2", item="Laser cutting:")
    excel_document.add_item(cell="P2", item="Nitrogen")
    excel_document.add_dropdown_selection(
        cell="P2", type="list", formula="'Sheet 2'!$A$3:$B$3"
    )
    excel_document.add_list(cell="B3", items=args[0], horizontal=False)  # File name
    excel_document.add_list(cell="C3", items=args[1], horizontal=False)  # Machine Time
    excel_document.add_list(cell="D3", items=args[2], horizontal=False)  # Weight
    excel_document.add_list(cell="E3", items=args[3], horizontal=False)  # Quantity

    num: int = 0

    for index in range(len(args[0])):
        row: int = num + 3
        excel_document.add_item(cell=f"F{row}", item=materials[0])  # Material Type
        excel_document.add_item(cell=f"G{row}", item=gauges[0])  # Gauge Selection
        excel_document.add_dropdown_selection(
            cell=f"F{row}", type="list", formula="'Sheet 2'!$A$1:$H$1"
        )
        excel_document.add_dropdown_selection(
            cell=f"G{row}", type="list", formula="'Sheet 2'!$A$2:$K$2"
        )

        cost_for_weight = f"INDEX('{path_to_sheet_prices}'!$D$6:$J$6,MATCH($F{row},'{path_to_sheet_prices}'!$D$5:$J$5,0))*$D{row}"
        cost_for_time = (
            f"(INDEX('Sheet 2'!$A$4:$B$4,MATCH($P$2,'Sheet 2'!$A$3:$B$3,0))/60)*$C{row}"
        )
        quantity = f"$E{row}"

        excel_document.add_item(
            cell=f"H{row}", item=f"=({cost_for_weight}+{cost_for_time})*{quantity}"
        )  # Cost

        for col in ["B", "C", "D", "E", "F", "G", "H"]:
            excel_document.set_alignment(
                cell=f"{col}{row}", horizontal="center", vertical="center", wrap_text=True
            )

        excel_document.format_cell(cell=f"H{row}", number_format="$#,##0.00")

        excel_document.add_image(
            cell=f"A{row}",
            path_to_image=f"{program_directory}/images/{args[4][index]}.jpeg",
        )
        excel_document.set_cell_height(cell=f"A{row}", height=size_of_picture / 1.3)

        num += 1

    excel_document.add_table(
        display_name="Table1", theme="TableStyleLight8", location=f"B2:H{num+2}"
    )

    excel_document.add_item(cell=f"G{num+4}", item="Total price: ")
    excel_document.add_item(
        cell=f"H{num+4}",
        item=f"=(SUM(Table1[Price ($)])/(1-($P${num+4})))*(1+$P${num+5})",
    )
    excel_document.format_cell(cell=f"H{num+4}", number_format="$#,##0.00")

    excel_document.add_item(cell=f"O{num + 4}", item="Overhead:")
    excel_document.add_item(cell=f"P{num + 4}", item=0.1)
    excel_document.format_cell(cell=f"P{num + 4}", number_format="0%")

    excel_document.add_item(cell=f"O{num + 5}", item="Markup:")
    excel_document.add_item(cell=f"P{num + 5}", item=0.5)
    excel_document.format_cell(cell=f"P{num + 5}", number_format="0%")

    excel_document.save()

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
        2 + (len(file_names) * 4),
        dual_line=True,
        title="Generating",
        force_tty=True,
        theme="smooth",
    ) as bar:
        part_dictionary = {}
        part_names = []
        quantity_numbers = []
        machining_times_numbers = []
        weights_numbers = []
        part_numbers = []
        extract_images_from_pdf(file_names, bar)
        for file_name in file_names:
            bar.text = "-> Getting text, please wait..."

            convert_pdf_to_text([file_name], bar)

            bar.text = "-> Getting images, please wait..."
            bar()

            bar.text = "-> Generating excel sheet, please wait..."
            bar()

            quantity_multiplier = get_table_value_from_text(regex=sheet_quantity_regex)[0]
            quantity_multiplier = int(
                quantity_multiplier.replace("PROGRAMME RUNS:  /  SCRAP: ", "")
            )

            part_file_paths = get_table_value_from_text(regex=geofile_name_regex)
            for pt in part_file_paths:
                pt = pt.split("\\")[-1].replace("\n", "").replace(".GEO", "")
                part_dictionary[pt] = {
                    "quantity": 0,
                    "machine_time": 0,
                    "weight": 0,
                    "part_number": 0,
                    "image_index": 0,
                }

                part_names.append(pt)
            # part_names.append([
            #     part_file_path.split("\\")[-1].replace("\n", "").replace(".GEO", "")
            #     for part_file_path in part_file_paths
            # ])

            quantity = get_table_value_from_text(regex=quantity_regex)
            # quantity_numbers.append([int(time.replace("  NUMBER: ", "")) for time in quantity])
            for q in quantity:
                q = q.replace("  NUMBER: ", "")
                quantity_numbers.append((int(q) * quantity_multiplier))

            machining_times = get_table_value_from_text(regex=machining_time_regex)
            for mt in machining_times:
                mt = mt.replace("MACHINING TIME: ", "").replace(" min", "")
                machining_times_numbers.append(float(mt))
            # machining_times_numbers.append([
            #     float(time.replace("MACHINING TIME: ", "").replace(" min", ""))
            #     for time in machining_times
            # ])

            weights = get_table_value_from_text(regex=weight_regex)
            for w in weights:
                w = w.replace("WEIGHT: ", "").replace(" lb", "")
                weights_numbers.append(float(w))
            # weights_numbers.append([
            #     float(time.replace("WEIGHT: ", "").replace(" lb", "")) for time in weights
            # ])

            part_numbers_string = get_table_value_from_text(regex=part_number_regex)
            for pn in part_numbers_string:
                pn = pn.replace("PART NUMBER: ", "")
                part_numbers.append(int(pn))
                # part_numbers.append([
                #     int(time.replace("PART NUMBER: ", "")) for time in part_numbers_string
                # ])

        for i, part_name in enumerate(part_names):
            part_dictionary[part_name]["quantity"] += quantity_numbers[i]
            part_dictionary[part_name]["machine_time"] = machining_times_numbers[i]
            part_dictionary[part_name]["weight"] = weights_numbers[i]
            part_dictionary[part_name]["part_number"] = part_numbers[i]
            part_dictionary[part_name]["image_index"] = i

        part_names = list(part_dictionary.keys())

        machining_times_numbers.clear()
        part_numbers.clear()
        quantity_numbers.clear()
        weights_numbers.clear()
        image_index = []

        for part in part_dictionary:
            machining_times_numbers.append(part_dictionary[part]["machine_time"])
            part_numbers.append(part_dictionary[part]["part_number"])
            quantity_numbers.append(part_dictionary[part]["quantity"])
            weights_numbers.append(part_dictionary[part]["weight"])
            image_index.append(part_dictionary[part]["image_index"])

        generate_excel_file(
            part_names,
            machining_times_numbers,
            weights_numbers,
            quantity_numbers,
            image_index,
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
