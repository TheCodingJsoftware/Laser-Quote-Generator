import re

import openpyxl
from openpyxl import Workbook
from openpyxl.drawing import image
from openpyxl.styles import Alignment
from openpyxl.utils.cell import column_index_from_string, get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo


class ExcelFile:
    """Create excel files easier with openpyxl"""

    def __init__(self, file_name: str) -> None:
        self.workbook = Workbook()
        self.worksheet = self.workbook.active

        self.cell_regex = r"^([A-Z]+)([1-9]\d*)$"
        self.file_name = file_name

    def parse_cell(self, cell: str) -> (str, int):
        """Parses excel cell input such as "AD300"

        Args:
            cell (str): input -> "AD300"

        Returns:
            str: "AD"
            int: 300
        """
        cell = cell.upper()
        if matches := re.search(self.cell_regex, cell):
            return (matches[1], int(matches[2]))

    def create_sheet(self, sheet_name: str) -> None:
        """Creates a new sheet within the excel work book

        Args:
            sheet_name (str): name of the sheet
        """
        self.workbook.create_sheet(sheet_name)

    def add_list_to_sheet(
        self, sheet_name: str, cell: str, items: list, horizontal: bool = True
    ) -> None:
        """Adds a list of items to the specfied sheet

        Args:
            sheet_name (str): Name of the sheet you want to add a list to.
            cell (str): specfied cell location, such as "A1"
            items (list): any list of items you want to add to the excel sheet
            horizontal (bool, optional): Allows for inputing lists vertical(False) or horizontal(True). Defaults to True.
        """
        col, row = self.parse_cell(cell=cell)
        sheet = self.workbook.get_sheet_by_name(sheet_name)
        if horizontal:
            col_index = column_index_from_string(col)
            for item in items:
                col_str = get_column_letter(col_index)
                try:
                    if item.is_integer():
                        sheet[f"{col_str}{row}"] = int(item)
                    elif not item.is_integer():
                        sheet[f"{col_str}{row}"] = float(item)
                except AttributeError:
                    sheet[f"{col_str}{row}"] = item
                col_index += 1
        else:
            for item in items:
                try:
                    if item.is_integer():
                        sheet[f"{col}{row}"] = int(item)
                    elif not item.is_integer():
                        sheet[f"{col}{row}"] = float(item)
                except AttributeError:
                    sheet[f"{col}{row}"] = item
                row += 1

    def add_list(self, cell: str, items: list, horizontal: bool = True) -> None:
        """Adds a list of items to the current workbook

        Args:
            sheet_name (str): Name of the sheet you want to add a list to.
            cell (str): specfied cell location, such as "A1"
            items (list): any list of items you want to add to the excel sheet
            horizontal (bool, optional): Allows for inputing lists vertical(False) or horizontal(True). Defaults to True.
        """
        col, row = self.parse_cell(cell=cell)
        if horizontal:
            col_index = column_index_from_string(col)
            for item in items:
                col_str = get_column_letter(col_index)
                try:
                    if item.is_integer():
                        self.worksheet[f"{col_str}{row}"] = int(item)
                    elif not item.is_integer():
                        self.worksheet[f"{col_str}{row}"] = float(item)
                except AttributeError:
                    self.worksheet[f"{col_str}{row}"] = item
                col_index += 1
        else:
            for item in items:
                try:
                    if item.is_integer():
                        self.worksheet[f"{col}{row}"] = int(item)
                    elif not item.is_integer():
                        self.worksheet[f"{col}{row}"] = float(item)
                except AttributeError:
                    self.worksheet[f"{col}{row}"] = item
                row += 1

    def add_item(self, cell: str, item) -> None:
        """Add any item to any cell in the excel work book

        Args:
            cell (str): Such as "A1"
            item (any): Any (item, str, int, float)
        """
        col, row = self.parse_cell(cell=cell)
        try:
            if item.is_integer():
                self.worksheet[f"{col}{row}"] = int(item)
            elif not item.is_integer():
                self.worksheet[f"{col}{row}"] = float(item)
        except AttributeError:
            self.worksheet[f"{col}{row}"] = item

    def set_cell_width(self, cell: str, width: int) -> None:
        """Change teh width of any cell, only the column is, the row is not used.

        Args:
            cell (str): Such as "A1"
            width (int): The width you want that column to be
        """
        col, _ = self.parse_cell(cell=cell)
        self.worksheet.column_dimensions[col].width = width

    def set_cell_height(self, cell: str, height: int) -> None:
        """Change teh width of any cell, only the row is, the column is not used.

        Args:
            cell (str): Such as "A1"
            height (int): The height you want that row to be
        """
        _, row = self.parse_cell(cell=cell)
        self.worksheet.row_dimensions[row].height = height

    def set_cell_size(self, cell: str, height: int, width: int) -> None:
        """Change the size of any cell

        Args:
            cell (str): Such as "A1
            height (int): The height you want that cell
            width (int): The width you want that cell
        """
        col, row = self.parse_cell(cell=cell)
        self.worksheet.column_dimensions[col].width = width
        self.worksheet.row_dimensions[row].height = height

    def add_image(self, cell: str, path_to_image: str) -> None:
        """Add an image to any cell

        Args:
            cell (str): Such as "A1"
            path_to_image (str): The direct path to the image
        """
        col, row = self.parse_cell(cell=cell)
        img = image.Image(path_to_image)
        img.anchor = f"{col}{row}"
        self.worksheet.add_image(img)

    def format_cell(self, cell: str, number_format: str) -> None:
        """Set the number format for any cell

        Args:
            cell (str): Such as "A1"
            number_format (str): The format you want, such as "$#,##0.00"
        """
        col, row = self.parse_cell(cell=cell)
        col_index = column_index_from_string(col)
        cell = self.worksheet.cell(row, col_index)
        cell.number_format = number_format

    def set_alignment(
        self, cell: str, horizontal: str, vertical: str, wrap_text: bool
    ) -> None:
        """Set the text alignment for any cell

        Args:
            cell (str): Such as "A1"
            horizontal (str): 'left', 'center', 'right'
            vertical (str): 'left', 'center', 'right'
            wrap_text (bool): True/False
        """
        col, row = self.parse_cell(cell=cell)
        col_index = column_index_from_string(col)
        cell = self.worksheet.cell(row, col_index)
        cell.alignment = Alignment(
            horizontal=horizontal, vertical=vertical, wrap_text=wrap_text
        )

    def add_dropdown_selection(self, cell: str, type: str, formula: str) -> None:
        """Add a data validation drop down selection for any cell

        Args:
            cell (str): Such as "A1"
            type (str): 'list'
            formula (str): the location of where the list is located such as: "A1:C1"
        """
        col, row = self.parse_cell(cell=cell)
        dropdown = DataValidation(type=type, formula1=formula)
        self.worksheet.add_data_validation(dropdown)
        dropdown.add(self.worksheet[f"${col}${row}"])

    def add_table(self, display_name: str, theme: str, location: str) -> None:
        """Add a table to the excel sheet

        Args:
            display_name (str): Name of that table, such as "Table1"
            theme (str): Any color theme provided by excel itself
            location (str): The location you want to format the table, such as: "A1:B3"
        """
        table = Table(displayName=display_name, ref=location)

        style = TableStyleInfo(
            name=theme,
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        table.autoFilter = None
        table.tableStyleInfo = style
        self.worksheet.add_table(table)

    def set_col_hidden(self, cell: str, hidden: bool = True) -> None:
        """Hide column

        Args:
            cell (str): Such as "A1"
            visible (bool): True or False
        """
        col, _ = self.parse_cell(cell=cell)
        self.worksheet.column_dimensions[col].hidden = hidden

    def set_row_hidden(self, cell: str, hidden: bool = True) -> None:
        """Hide row

        Args:
            cell (str): Such as "A1"
            visible (bool): True or False
        """
        _, row = self.parse_cell(cell=cell)
        self.worksheet.row_dimensions[col].hidden = hidden

    def save(self) -> None:
        """Save excel file."""
        self.workbook.save(self.file_name)
