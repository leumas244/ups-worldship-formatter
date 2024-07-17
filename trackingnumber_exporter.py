import xml.etree.ElementTree as ElementTree
import os
import re

import settings
from data_classes import Package
import excel_converter

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
        return 'no_file'
        


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


def start_routine() -> None:
    excel_and_out_files = get_matching_excel_and_out_files()
    
    for excel_file_path, out_file_path in excel_and_out_files:
        xml_root = get_xml_tree(out_file_path)
        
        proccesed_packages: list[Package] = []
        
        openShipments = xml_root.findall("{x-schema:OpenShipments.xdr}OpenShipment")
        for openShipment in openShipments:
            process_status = openShipment.get("ProcessStatus")
            if process_status == "Processed":
                shipTo = openShipment.find("{x-schema:OpenShipments.xdr}ShipTo")
                shipmentInformation = openShipment.find("{x-schema:OpenShipments.xdr}ShipmentInformation")
                package = Package(excelReciverString="", excel_row=0, excel_column=0)
                package.recipientName = shipTo.find("{x-schema:OpenShipments.xdr}CompanyOrName").text
                package.packageCount = int(shipmentInformation.find("{x-schema:OpenShipments.xdr}NumberOfPackages").text)
                
                trackingNumbers = openShipment.find("{x-schema:OpenShipments.xdr}ProcessMessage").find("{x-schema:OpenShipments.xdr}TrackingNumbers")
                for trackingNumber in trackingNumbers:
                    package.trackingNumbers.append(trackingNumber.text)
                    
                proccesed_packages.append(package)
        
        
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
            last_reciverString: str = "set"
            for row in range(titleRow + 1, 200):
                if sheet.cell(row=row, column=reciverColum).value:
                    if not sheet.cell(row=row, column=trackingNumberColum).value:
                        reciverString = sheet.cell(row=row, column=reciverColum).value
                        
                        for package in proccesed_packages:
                            if package.recipientName in reciverString:
                                if len(package.trackingNumbers) == 1:
                                    if len(package.trackingNumbers) == sheet.cell(row=row, column=packageCountColum).value:
                                        sheet.cell(row=row, column=shippingServiceColum).value = "UPS"
                                        sheet.cell(row=row, column=trackingNumberColum).value = package.trackingNumbers[0]
                                elif len(package.trackingNumbers) > 1:
                                    if len(package.trackingNumbers) == sheet.cell(row=row, column=packageCountColum).value:
                                        sheet.cell(row=row, column=shippingServiceColum).value = "UPS"
                                        out = ''
                                        for trackingNumber in package.trackingNumbers:
                                            out = out + trackingNumber + ', '
                                        sheet.cell(row=row, column=trackingNumberColum).value = out
                                        # ToDo: was Wenn ein paket in mehrer zeilen aufgeteilt wurde


                else:
                    if sheet.cell(row=row, column=referenceColum).value:
                        continue

                    else:
                        break
                last_reciverString = sheet.cell(row=row, column=reciverColum).value
            workbook.save(excel_file_path)