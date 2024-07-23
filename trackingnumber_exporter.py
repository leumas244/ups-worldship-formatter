from datetime import datetime
import shutil
import xml.etree.ElementTree as ElementTree
import os
import re

import settings
from data_classes import Package
import excel_converter


def print_info(info: str, tab: int = 0, time_stamp: bool = True) -> None:
    columns, rows = shutil.get_terminal_size()
    columns -= 1
    pre_print_info = get_pre_print_info() + ("   " * tab)
    if time_stamp:
        print(pre_print_info, end="")
    else:
        print(" " * len(pre_print_info), end="")
    if (len(pre_print_info) + len(info)) > columns:
        result = split_string_by_length(info, (columns - len(pre_print_info)))
        for part in result:
            if result.index(part) == 0:
                print(part)
            else:
                print((" " * len(pre_print_info)) + part)
    else:
        print(info)


def get_pre_print_info() -> str:
    now_variable = datetime.now()
    pre_print_info = f"[{now_variable.strftime('%d/%m/%Y %H:%M:%S')}] "
    return pre_print_info


def split_string_by_length(s: str, length: int) -> list[str]:
    return [s[i : i + length] for i in range(0, len(s), length)]


def get_basename_from_file_path(path_to_file: str) -> str:
    return os.path.basename(path_to_file)


def get_all_processed_xml_files() -> list[str]:
    files = []
    out_pattern = r"\.Out$"
    try:
        for file in os.listdir(settings.xml_output_file_folder):
            if re.search(out_pattern, file):
                file_name = os.path.join(settings.xml_output_file_folder, file)
                if os.path.isfile(file_name):
                    files.append(file_name)

        return files

    except Exception as e:
        # print_info(f"Fehler beim Sammeln der zu analysierdenden Excel Dateien: {e}")
        # print_info(f"Das Programm beendet sich mit Problemen")
        exit()


def get_all_excel_folder() -> list[str]:
    folders = []
    try:
        for folder in os.listdir(settings.parsed_excel_file_folder):
            folder_name = os.path.join(settings.parsed_excel_file_folder, folder)
            if os.path.isdir(folder_name):
                folders.append(folder_name)
        return folders

    except Exception as e:
        # print_info(f"Fehler beim Sammeln der zu analysierdenden Excel Dateien: {e}")
        # print_info(f"Das Programm beendet sich mit Problemen")
        exit()


def get_all_excel_file_in_folder(excel_folder: str) -> str:
    found_excel_files = []
    xlsx_pattern = r"\.xlsx$"
    beginning_pattern = r"^~\$"
    for file in os.listdir(excel_folder):
        if re.search(xlsx_pattern, file):
            if re.search(beginning_pattern, file):
                return "file_is_used"

    for file in os.listdir(excel_folder):
        if re.search(xlsx_pattern, file):
            file_name = os.path.join(excel_folder, file)
            found_excel_files.append(file_name)

    if len(found_excel_files) == 1:
        return found_excel_files[0]
    else:
        return "no_file"


def get_matching_excel_and_out_files() -> list[tuple[str, str]]:
    matching_excel_and_out_files = []
    out_files = get_all_processed_xml_files()
    excel_folders = get_all_excel_folder()

    for excel_folder in excel_folders:
        folder_name = get_basename_from_file_path(excel_folder)
        for out_file_path in out_files:
            out_file_name = get_basename_from_file_path(out_file_path)
            if folder_name in out_file_name:
                excel_file_path = get_all_excel_file_in_folder(excel_folder)
                matching = (excel_file_path, out_file_path)
                matching_excel_and_out_files.append(matching)
                out_files.remove(out_file_path)
                break

    return matching_excel_and_out_files


def get_xml_tree(file_path: str) -> ElementTree:
    tree = ElementTree.parse(file_path)
    return tree.getroot()


def create_trackingNumbers_and_refNumbers_assignment(package: Package) -> Package:

    if len(package.referenceNumbers) > 1:
        # ToDo Zeile manuell Prüfen lassen
        if len(package.referenceNumbers) == len(package.trackingNumbers):
            referenceNumbers = []
            for referenceNumber in package.referenceNumbers:
                newReferenceTuple = (referenceNumber[0], 1)
                referenceNumbers.append(newReferenceTuple)
            package.referenceNumbers = referenceNumbers

            trackingNumberCounter = 0
            for referenceNumber in package.referenceNumbers:
                package.excelTrackingAssignment["name"] = package.recipientName
                package.excelTrackingAssignment["referenceNumber"] = referenceNumber[0]
                package.excelTrackingAssignment["packageCount"] = referenceNumber[1]
                package.excelTrackingAssignment["trackingNumbers"] = [
                    package.trackingNumbers[trackingNumberCounter]
                ]
                trackingNumberCounter += 1
        else:
            # ToDo: Fehler melden, dass die Nummern Selbst eingetragen werden müssen
            pass

    elif len(package.referenceNumbers) == 1:
        package.excelTrackingAssignment["name"] = package.recipientName
        package.excelTrackingAssignment["referenceNumber"] = package.referenceNumbers[
            0
        ][0]
        package.excelTrackingAssignment["packageCount"] = package.referenceNumbers[0][1]
        package.excelTrackingAssignment["trackingNumbers"] = package.trackingNumbers
    else:
        # ToDo: Fehler keine Reference Nummer Gefunden
        pass

    return package


def get_proccesed_packages(xml_root: ElementTree) -> list[Package]:
    proccesed_packages: list[Package] = []

    openShipments = xml_root.findall("{x-schema:OpenShipments.xdr}OpenShipment")
    for openShipment in openShipments:
        process_status = openShipment.get("ProcessStatus")
        if process_status == "Processed":
            shipTo = openShipment.find("{x-schema:OpenShipments.xdr}ShipTo")
            shipmentInformation = openShipment.find(
                "{x-schema:OpenShipments.xdr}ShipmentInformation"
            )
            processMessage = openShipment.find(
                "{x-schema:OpenShipments.xdr}ProcessMessage"
            )
            package = Package(
                excelReciverString="", excel_row=0, excel_column=0, coordinate=""
            )
            package.recipientName = shipTo.find(
                "{x-schema:OpenShipments.xdr}CompanyOrName"
            ).text
            package.address1 = shipTo.find("{x-schema:OpenShipments.xdr}Address1").text
            package.postalCode = shipTo.find(
                "{x-schema:OpenShipments.xdr}PostalCode"
            ).text
            package.packageCount = int(
                shipmentInformation.find(
                    "{x-schema:OpenShipments.xdr}NumberOfPackages"
                ).text
            )

            for count in range(1, 6):
                ref_number_tree = processMessage.find(
                    ("{x-schema:OpenShipments.xdr}Reference" + str(count))
                )
                if ref_number_tree is not None:
                    if ref_number_tree.text:
                        refrence_tuple: tuple = (
                            ref_number_tree.text,
                            package.packageCount,
                        )
                        package.referenceNumbers.append(refrence_tuple)

            trackingNumbers = processMessage.find(
                "{x-schema:OpenShipments.xdr}TrackingNumbers"
            )
            for trackingNumber in trackingNumbers:
                package.trackingNumbers.append(trackingNumber.text)

            package = create_trackingNumbers_and_refNumbers_assignment(package)

            proccesed_packages.append(package)
    return proccesed_packages


def detect_packages_from_the_same_recipient(
    proccesed_packages: list[Package],
) -> list[Package]:
    detected_packages: list[Package] = []
    for proccesed_package in proccesed_packages:
        is_in_detected_packages = False
        for detected_package in detected_packages:
            if (
                proccesed_package.recipientName == detected_package.recipientName
                and proccesed_package.address1 == detected_package.address1
                and proccesed_package.postalCode == detected_package.postalCode
            ):
                if (
                    len(proccesed_package.referenceNumbers) == 1
                    and len(detected_package.referenceNumbers) == 1
                ):
                    if (
                        proccesed_package.referenceNumbers[0]
                        == detected_package.referenceNumbers[0]
                    ):
                        is_in_detected_packages = True
                        new_detected_package = detected_package
                        index_detected_package = detected_packages.index(
                            detected_package
                        )
                        new_detected_package.packageCount += (
                            proccesed_package.packageCount
                        )
                        new_refrence_tuple = (
                            detected_package.referenceNumbers[0][0],
                            new_detected_package.packageCount,
                        )
                        new_detected_package.referenceNumbers = [new_refrence_tuple]
                        for trackingNumber in proccesed_package.trackingNumbers:
                            new_detected_package.trackingNumbers.append(trackingNumber)

                        new_detected_package.excelTrackingAssignment["packageCount"] = (
                            new_detected_package.referenceNumbers[0][1]
                        )
                        new_detected_package.excelTrackingAssignment[
                            "trackingNumbers"
                        ] = new_detected_package.trackingNumbers

                        detected_packages[index_detected_package] = new_detected_package

                elif len(proccesed_package.referenceNumbers) > 1:
                    # ToDo: Fehler werfen
                    is_in_detected_packages = True
                    pass

        if not is_in_detected_packages:
            detected_packages.append(proccesed_package)

    return detected_packages
        

def get_merged_cell_value(sheet_, row, column):
    # Überprüfen Sie, ob die Zelle Teil einer zusammengeführten Zelle ist
    for merged_cell_range in sheet_.merged_cells.ranges:
        column_row_tuple = (row, column)
        if column_row_tuple in merged_cell_range.left:
            # Wenn ja, finden Sie die Zelle, die den Wert enthält
            return merged_cell_range.start_cell.value
    # Wenn die Zelle nicht Teil einer zusammengeführten Zelle ist, geben Sie den Zellenwert zurück
    return sheet_.cell(row=row, column=column).value


def start_routine() -> None:
    print_info(f"Starte das Importieren der Tracking-Nummern in Excel-Files")
    matching_excel_and_out_files = get_matching_excel_and_out_files()

    for excel_file_path, out_file_path in matching_excel_and_out_files:
        print_info(f"Bearbeite die Datei '{get_basename_from_file_path(excel_file_path)}'")
        xml_root = get_xml_tree(out_file_path)

        proccesed_packages = get_proccesed_packages(xml_root)

        formed_packages = detect_packages_from_the_same_recipient(proccesed_packages)

        workbook = excel_converter.get_workbook(excel_file_path)
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

        if (
            titleRow
            and senderColum
            and reciverColum
            and referenceColum
            and packageCountColum
            and shippingServiceColum
            and trackingNumberColum
        ):
            for row in range(titleRow + 1, 200):
                if sheet.cell(row=row, column=reciverColum).value:
                    if not sheet.cell(row=row, column=trackingNumberColum).value:
                        reciverString = sheet.cell(row=row, column=reciverColum).value

                        for package in formed_packages:
                            if package.recipientName in reciverString:
                                if len(package.trackingNumbers) == 1:
                                    if (
                                        len(package.trackingNumbers)
                                        == sheet.cell(
                                            row=row, column=packageCountColum
                                        ).value
                                        and package.referenceNumbers[0][0]
                                        == sheet.cell(
                                            row=row, column=referenceColum
                                        ).value
                                    ):
                                        sheet.cell(
                                            row=row, column=shippingServiceColum
                                        ).value = "UPS"
                                        sheet.cell(
                                            row=row, column=trackingNumberColum
                                        ).value = package.trackingNumbers[0]
                                        formed_packages.remove(package)
                                        break
                                elif len(package.trackingNumbers) > 1:
                                    if (
                                        len(package.trackingNumbers)
                                        == sheet.cell(
                                            row=row, column=packageCountColum
                                        ).value
                                        and package.referenceNumbers[0][0]
                                        == sheet.cell(
                                            row=row, column=referenceColum
                                        ).value
                                    ):
                                        sheet.cell(
                                            row=row, column=shippingServiceColum
                                        ).value = "UPS"
                                        out = ""
                                        for trackingNumber in package.trackingNumbers:
                                            out = out + trackingNumber + ", "
                                        sheet.cell(
                                            row=row, column=trackingNumberColum
                                        ).value = out
                                        formed_packages.remove(package)
                                        break

                else:
                    if sheet.cell(row=row, column=referenceColum).value:
                        if not sheet.cell(row=row, column=trackingNumberColum).value:
                            reciverString = get_merged_cell_value(sheet_=sheet, row=row, column=reciverColum)

                            for package in formed_packages:
                                if package.recipientName in reciverString:
                                    if len(package.trackingNumbers) == 1:
                                        if (
                                            len(package.trackingNumbers)
                                            == sheet.cell(
                                                row=row, column=packageCountColum
                                            ).value
                                            and package.referenceNumbers[0][0]
                                            == sheet.cell(
                                                row=row, column=referenceColum
                                            ).value
                                        ):
                                            sheet.cell(
                                                row=row, column=shippingServiceColum
                                            ).value = "UPS"
                                            sheet.cell(
                                                row=row, column=trackingNumberColum
                                            ).value = package.trackingNumbers[0]
                                            formed_packages.remove(package)
                                            break
                                    elif len(package.trackingNumbers) > 1:
                                        if (
                                            len(package.trackingNumbers)
                                            == sheet.cell(
                                                row=row, column=packageCountColum
                                            ).value
                                            and package.referenceNumbers[0][0]
                                            == sheet.cell(
                                                row=row, column=referenceColum
                                            ).value
                                        ):
                                            sheet.cell(
                                                row=row, column=shippingServiceColum
                                            ).value = "UPS"
                                            out = ""
                                            for (
                                                trackingNumber
                                            ) in package.trackingNumbers:
                                                out = out + trackingNumber + ", "
                                            sheet.cell(
                                                row=row, column=trackingNumberColum
                                            ).value = out
                                            formed_packages.remove(package)
                                            break

                    else:
                        break

            for row in range(titleRow + 1, 200):
                if sheet.cell(row=row, column=reciverColum).value:
                    if not sheet.cell(row=row, column=trackingNumberColum).value:
                        print_info(f"- Bitte Zeile {row} selbst eintragen", tab=1)
                else:
                    if sheet.cell(row=row, column=referenceColum).value:
                        if not sheet.cell(row=row, column=trackingNumberColum).value:
                            print_info(f"- Bitte Zeile {row} selbst eintragen", tab=1)
                    else:
                        break

            workbook.save(excel_file_path)
            
    print_info(f"Beende das Importieren der Tracking-Nummern in Excel-Files")


start_routine()
