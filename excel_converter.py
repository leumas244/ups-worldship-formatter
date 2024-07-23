import openpyxl
import warnings

from data_classes import Package

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


def get_workbook(excel_file_name: str) -> openpyxl.Workbook:
    return openpyxl.load_workbook(filename=excel_file_name)


def get_packages_from_excel_file(excel_file: str) -> list[Package]:
    workbook = get_workbook(excel_file)
    sheet = workbook.active

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
