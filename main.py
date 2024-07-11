import os
import shutil
from datetime import datetime

import settings
import excel_converter
import address_parser
from data_classes import Package
import export_manager


def print_info(info: str) -> None:
    columns, rows = shutil.get_terminal_size()
    columns -= 1
    now_variable = datetime.now()
    pre_print_info = f"[main_process | {now_variable.strftime('%d/%m/%Y %H:%M:%S')}] "
    print(pre_print_info, end="")
    if (len(pre_print_info) + len(info)) > columns:
        result = split_string_by_length(info, (columns - len(pre_print_info)))
        for part in result:
            if result.index(part) == 0:
                print(part)
            else:
                print((" " * len(pre_print_info)) + part)
    else:
        print(info)


def split_string_by_length(s, length):
    return [s[i : i + length] for i in range(0, len(s), length)]


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


def has_package_all_needed_informations(package: Package) -> bool:
    if not package.recipientName:
        print_info("Kein Name")
        return False
    elif not package.address1:
        print_info("Keine Adresse")
        return False
    elif not package.country:
        print_info("Kein Land")
        return False
    elif not package.postalCode:
        print_info("Keine PLZ")
        return False
    elif not package.city:
        print_info("Keine Stadt")
        return False
    elif not package.referenceNumber:
        print_info("Keine Referenznummer")
        return False
    else:
        return True


def int_to_alphabet(num):
    """Konvertiert eine Zahl in einen entsprechenden Buchstaben oder eine Buchstabenfolge nach dem Alphabet."""
    result = ""
    while num > 0:
        num -= 1  # Um den Index bei 0 zu beginnen (A = 0, B = 1, ...)
        result = chr(num % 26 + 65) + result
        num //= 26
    return result


def ckeck_package_on_abroad_and_dublicate(package: Package) -> list[Package]:
    output_list = []
    if package.country != "DE" and package.packageCount > 1:
        packageCount = package.packageCount
        package = set_package_to_single_package(package)
        output_list.append(package)
        for i in range(1, packageCount):
            dublicate_of_package = dublicate_package(package)
            output_list.append(dublicate_of_package)
        return output_list
    else:
        return [package]


def set_package_to_single_package(package: Package) -> Package:
    package.packageCount = 1
    package.weight = 10.0
    return package


def dublicate_package(package: Package) -> Package:
    new_package = Package(
        excelReciverString=package.excelReciverString,
        excel_row=package.excel_row,
        excel_column=package.excel_column,
    )
    new_package.recipientName = package.recipientName
    new_package.recipientNameAddtional = package.recipientNameAddtional
    new_package.address1 = package.address1
    new_package.address2 = package.address2
    new_package.address3 = package.address3
    new_package.country = package.country
    new_package.postalCode = package.postalCode
    new_package.city = package.city
    new_package.state = package.state
    new_package.phoneNumber = package.phoneNumber
    new_package.email = package.email
    new_package.weight = package.weight
    new_package.service = package.service
    new_package.referenceNumber = package.referenceNumber
    new_package.packageCount = package.packageCount
    return new_package


def main() -> None:
    print_info(f"Das Programm startet.")
    inital_check_on_existing_file_infrastructure()
    print()

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
                    f"Bite gib die Adresse in der Datei '{get_file_name_from_file_path(excel_file)}' in Zelle {int_to_alphabet(package.excel_column)}{package.excel_row} manuel ein!"
                )
                continue

            try:
                package = address_parser.sort_assignment_to_package(
                    address_assignment, package
                )
            except Exception as e:
                excel_file_has_a_problem = True
                print_info({str(e)})
                print_info(
                    f"Bite gib die Adresse in der Datei '{get_file_name_from_file_path(excel_file)}' in Zelle {int_to_alphabet(package.excel_column)}{package.excel_row} manuel ein!"
                )
                continue

            package_state = has_package_all_needed_informations(package)

            additonal_packages_for_abroads = ckeck_package_on_abroad_and_dublicate(
                package
            )

            if package_state:
                for ready_packages in additonal_packages_for_abroads:
                    output_packages.append(ready_packages)
            else:
                excel_file_has_a_problem = True
                print_info(
                    f"Bite gib die Adresse in der Datei '{get_file_name_from_file_path(excel_file)}' in Zelle {int_to_alphabet(package.excel_column)}{package.excel_row} manuel ein!"
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
