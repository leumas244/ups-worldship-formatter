import os
import shutil
from datetime import datetime
import xml.etree.ElementTree as ET

import settings
import excel_converter
import address_parser
from data_classes import Package
import export_manager


def print_info(info: str) -> None:
    now_variable = datetime.now()
    print(f"[main_process | {now_variable.strftime('%d/%m/%Y %H:%M:%S')}] {info}")


def get_file_name_from_file_path(path_to_file: str) -> str:
    return os.path.basename(path_to_file)


def move_file(source_file: str, destination_folder: str):
    try:
        file_name = os.path.basename(source_file)

        destination_file = os.path.join(destination_folder, file_name)

        shutil.move(source_file, destination_file)
        print_info(
            f"Die Datei '{file_name}' wurde erfolgreich nach '{destination_folder}' verschoben."
        )

    except Exception as e:
        print_info(f"Fehler beim Verschieben der Datei: {e}")
        print_info(
            f"Das Programm arbeitet weiter, aber die ausgewerteten Excel-Dateien wurden nicht verschoben!"
        )


def inital_check_on_existing_file_infrastructure() -> None:
    for folder in settings.all_needed_folders:
        if not os.path.exists(folder):
            try:
                os.makedirs(folder)
                print_info(f"Das Verzeichnis '{folder}' wurde erstellt.")
            except Exception as e:
                print_info(f"Fehler beim Erstellen des Verzeichnisses '{folder}': {e}")
                print_info(f"Das Programm beendet sich mit Problemen")
                exit()
    return


def get_files_to_parse() -> list[str]:
    files = []
    try:
        for file in os.listdir(settings.not_parsed_excel_file_folder):
            file_name = os.path.join(settings.not_parsed_excel_file_folder, file)
            if os.path.isfile(file_name):
                files.append(file_name)
        return files

    except Exception as e:
        print_info(f"Fehler beim Sammeln der zu analysierdenden Excel Dateien: {e}")
        print_info(f"Das Programm beendet sich mit Problemen")
        exit()


def write_packages_to_xml_file(packages: list[Package], excel_file: str) -> str:
    now_variable = datetime.now()
    file_name = f"-{get_file_name_from_file_path(excel_file).replace('.xlsx', '')}-parsed_at_{now_variable.strftime('%d.%m.%Y_%H-%M-%S')}.xml"
    file_path = os.path.join(settings.csv_output_file_folder, file_name)
    xml_tree = export_manager.get_xml_tree(packages)
    xml_tree.write(file_path, encoding="utf-8", xml_declaration=True)
    return file_name


def main() -> None:
    print_info(f"Das Programm startet.")
    inital_check_on_existing_file_infrastructure()

    excel_files_to_parse = get_files_to_parse()

    for excel_file in excel_files_to_parse:
        excel_file_has_a_problem = False
        try:
            packages = excel_converter.get_packages_from_excel_file(excel_file)
        except Exception as e:
            print_info(f"Fehler beim Auswerten der Excel-Datei '{excel_file}': {e}")
            print_info(f"Das Datei wird übersprungen!")
            continue

        output_packages: list[Package] = []
        for package in packages:
            try:
                address_assignment = address_parser.parse_address(
                    package.excelReciverString
                )
            except Exception as e:
                excel_file_has_a_problem = True
                print_info({str(e)})
                print_info(
                    f"Bite gib die Adresse in der Datei '{get_file_name_from_file_path(excel_file)}' in Zeile {package.excel_row} manuel ein!"
                )
                continue

            try:
                package = address_parser.sort_assignment_to_package(
                    address_assignment, package
                )
                output_packages.append(package)
            except Exception as e:
                excel_file_has_a_problem = True
                print_info({str(e)})
                print_info(
                    f"Bite gib die Adresse in der Datei '{get_file_name_from_file_path(excel_file)}' in Zeile {package.excel_row} manuel ein!"
                )
                continue

            output_file_name = write_packages_to_xml_file(output_packages, excel_file)

        if excel_file_has_a_problem:
            move_file(excel_file, settings.parsed_excel_file_with_problems_folder)
        else:
            move_file(excel_file, settings.parsed_excel_file_folder)

    print_info(f"Das Programm ist beendet.")
    print()


main()
