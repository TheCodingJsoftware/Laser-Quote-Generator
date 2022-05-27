import openpyxl
from openpyxl import Workbook
from openpyxl.drawing import image
from openpyxl.styles import Alignment
from openpyxl.utils.cell import column_index_from_string, get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo


class ExcelFile:
    def __init__(self, file_name: str) -> None:
        self.workbook = Workbook()
        self.worksheet = self.workbook.active

        self.file_name = file_name

    def create_sheet(self, sheet_name: str) -> None:
        self.workbook.create_sheet(sheet_name)

    def add_list_to_sheet(
        self, sheet_name: str, col: str, row: int, items: list, horizontal: bool = True
    ) -> None:
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

    def add_list(self, col: str, row: int, items: list, horizontal: bool = True) -> None:
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

    def add_item(self, col: str, row: int, item) -> None:
        try:
            if item.is_integer():
                self.worksheet[f"{col}{row}"] = int(item)
            elif not item.is_integer():
                self.worksheet[f"{col}{row}"] = float(item)
        except AttributeError:
            self.worksheet[f"{col}{row}"] = item

    def set_col_width(self, col: str, width: int) -> None:
        self.worksheet.column_dimensions[col].width = width

    def set_row_height(self, row: int, height: int) -> None:
        self.worksheet.row_dimensions[row].height = height

    def set_cell_size(self, col: str, row: int, height: int, width: int) -> None:
        self.worksheet.column_dimensions[col].width = width
        self.worksheet.row_dimensions[row].height = height

    def add_image(self, col: str, row: int, path_to_image: str) -> None:
        img = image.Image(path_to_image)
        img.anchor = f"{col}{row}"
        self.worksheet.add_image(img)

    def format_cell(self, col: str, row: int, number_format: str) -> None:
        col_index = column_index_from_string(col)
        cell = self.worksheet.cell(row, col_index)
        cell.number_format = number_format

    def set_alignment(
        self, col: str, row: int, horizontal: str, vertical: str, wrap_text: bool
    ) -> None:
        col_index = column_index_from_string(col)
        cell = self.worksheet.cell(row, col_index)
        cell.alignment = Alignment(
            horizontal=horizontal, vertical=vertical, wrap_text=wrap_text
        )

    def add_dropdown_selection(self, col: str, row: int, type: str, formula: str) -> None:
        dropdown = DataValidation(type=type, formula1=formula)
        self.worksheet.add_data_validation(dropdown)
        dropdown.add(self.worksheet[f"${col}${row}"])

    def add_table(self, display_name: str, theme: str, location: str) -> None:
        table = Table(displayName=display_name, ref=location)

        style = TableStyleInfo(
            name=theme,
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        table.tableStyleInfo = style
        self.worksheet.add_table(table)

    def save(self) -> None:
        self.workbook.save(self.file_name)
