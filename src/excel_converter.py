import openpyxl
import warnings

import openpyxl.cell
import openpyxl.worksheet
import openpyxl.worksheet.merge
import openpyxl.worksheet.worksheet

from data_classes import Package

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


def get_workbook(excel_file_name: str) -> openpyxl.Workbook:
    return openpyxl.load_workbook(filename=excel_file_name)


def get_active_sheet_from_excel_file(
    excel_file_name: str,
) -> openpyxl.worksheet.worksheet.Worksheet:
    workbook = get_workbook(excel_file_name)
    return workbook.active


def cell_is_in_merge_cell_range(
    merged_cell_range: openpyxl.worksheet.merge.MergedCellRange,
    cell: openpyxl.cell.Cell,
) -> bool:
    min_row = merged_cell_range.min_row
    min_col = merged_cell_range.min_col
    max_row = merged_cell_range.max_row
    max_col = merged_cell_range.max_col

    if (min_row <= cell.row <= max_row) and (min_col <= cell.column <= max_col):
        return True
    else:
        return False


def is_cell_part_of_merged_cell(
    sheet: openpyxl.worksheet.worksheet.Worksheet, cell: openpyxl.cell.Cell
) -> bool:
    for merged_cell_range in sheet.merged_cells.ranges:
        if cell_is_in_merge_cell_range(merged_cell_range, cell):
            return True
    return False


def get_merged_cell_value(
    sheet: openpyxl.worksheet.worksheet.Worksheet, cell: openpyxl.cell.Cell
):
    if isinstance(cell, openpyxl.cell.cell.MergedCell):
        merged_cell_range: openpyxl.worksheet.merge.MergedCellRange
        for merged_cell_range in sheet.merged_cells.ranges:
            if cell_is_in_merge_cell_range(merged_cell_range, cell):
                start_cell: openpyxl.cell.Cell = merged_cell_range.start_cell
                return start_cell.value
    else:
        return cell.value
    

def get_excel_type_from_reciverColum(sheet: openpyxl.worksheet.worksheet.Worksheet, reciverColumn: int):
    pass


def get_type_and_headerCells_from_excelSheet(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
) -> dict[str, int]:
    sheet_informations = {
        "excel_sheet_type": None,
        "titleRow": None,
        "senderColum": None,
        "reciverColum": None,
        "referenceColum": None,
        "packageCountColum": None,
        "shippingServiceColum": None,
        "trackingNumberColum": None,
    }

    for rowCounter in range(1, 5):
        if not sheet_informations["titleRow"]:
            for columCounter in range(1, 8):
                cell = sheet.cell(row=rowCounter, column=columCounter)
                if cell.value == "Sender":
                    sheet_informations["titleRow"] = rowCounter
                    sheet_informations["senderColum"] = columCounter
                elif cell.value == "Empfänger":
                    sheet_informations["titleRow"] = rowCounter
                    sheet_informations["reciverColum"] = columCounter
                elif cell.value == "Variante / Farbe":
                    sheet_informations["referenceColum"] = columCounter
                elif cell.value == "Menge":
                    sheet_informations["packageCountColum"] = columCounter
                elif cell.value == "Versand-Dienstleister":
                    sheet_informations["shippingServiceColum"] = columCounter
                elif cell.value == "Sendungs-Nummer":
                    sheet_informations["trackingNumberColum"] = columCounter
        else:
            break

    excel_type = get_excel_type_from_reciverColum(sheet_informations["reciverColum"])


def get_packages_from_excel_file(excel_file_name: str) -> list[Package]:
    sheet = get_active_sheet_from_excel_file(excel_file_name)

    titleRow = None
    senderColum = None
    reciverColum = None
    referenceColum = None
    packageCountColum = None
    shippingServiceColum = None
    trackingNumberColum = None

    for rowCounter in range(1, 4):
        if not titleRow:
            for columCounter in range(1, 7):
                cell = sheet.cell(row=rowCounter, column=columCounter)
                if cell.value == "Sender":
                    titleRow = rowCounter
                    senderColum = columCounter
                elif cell.value == "Empfänger":
                    titleRow = rowCounter
                    reciverColum = columCounter
                elif cell.value == "Variante / Farbe":
                    referenceColum = columCounter
                elif cell.value == "Menge":
                    packageCountColum = columCounter
                elif cell.value == "Versand-Dienstleister":
                    shippingServiceColum = columCounter
                elif cell.value == "Sendungs-Nummer":
                    trackingNumberColum = columCounter

    packages: list[Package] = []

    if (
        titleRow
        and senderColum
        and reciverColum
        and referenceColum
        and packageCountColum
    ):
        for row in range(titleRow + 1, 200):
            if sheet.cell(row=row, column=reciverColum).value:
                reciverString = sheet.cell(row=row, column=reciverColum).value
                coordinate = sheet.cell(row=row, column=reciverColum).coordinate
                newPackage = Package(
                    excelReciverString=reciverString,
                    excel_row=row,
                    excel_column=reciverColum,
                    coordinate=coordinate,
                )
                refrenceTuple: tuple = (
                    sheet.cell(row=row, column=referenceColum).value,
                    sheet.cell(row=row, column=packageCountColum).value,
                )
                newPackage.referenceNumbers.append(refrenceTuple)
                newPackage.packageCount = int(
                    sheet.cell(row=row, column=packageCountColum).value
                )

                packages.append(newPackage)
            else:
                if sheet.cell(row=row, column=referenceColum).value:
                    packages.remove(newPackage)
                    refrenceTuple: tuple = (
                        sheet.cell(row=row, column=referenceColum).value,
                        sheet.cell(row=row, column=packageCountColum).value,
                    )
                    newPackage.referenceNumbers.append(refrenceTuple)
                    newPackage.packageCount = newPackage.packageCount + (
                        int(sheet.cell(row=row, column=packageCountColum).value)
                    )
                    packages.append(newPackage)

                else:
                    break

    return packages
