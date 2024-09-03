import os
import re
import shutil
from datetime import datetime

import settings
import excel_converter
import address_parser
from data_classes import Package
import export_manager


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


def get_file_name_from_file_path(path_to_file: str) -> str:
    return os.path.basename(path_to_file)


def print_excel_file_info(excel_files_to_parse: list[str]) -> None:
    pre_print_info = get_pre_print_info()
    info = "Gefundene Datei(en) zum analysieren:"
    print(pre_print_info + info)
    for excel_file in excel_files_to_parse:
        file_name = get_file_name_from_file_path(excel_file)
        print((" " * len(pre_print_info)) + "- " + file_name)
    print()


def move_file(source_file: str) -> str:
    folder_for_file = settings.parsed_excel_file_folder + get_file_name_from_file_path(source_file).replace(".xlsx", "/")
    try:
        if not os.path.exists(folder_for_file):
            os.makedirs(folder_for_file)
    except:
        print_info(f"Fehler beim erstellen des Ordners '{folder_for_file}'", 1)
        return
    try:
        file_name = os.path.basename(source_file)

        destination_file = os.path.join(folder_for_file, file_name)

        shutil.move(source_file, destination_file)
        print_info(
            f"Die Datei '{file_name}' wurde erfolgreich nach '{folder_for_file}' verschoben.",
            1,
        )
        return folder_for_file

    except Exception as e:
        print_info(f"Fehler beim Verschieben der Datei: {e}", 1)
        print_info(
            f"Das Programm arbeitet weiter, aber die ausgewerteten Excel-Dateien wurden nicht verschoben!",
            1,
        )
        return
        
        
def create_text_file_with_problem_information(folder_for_file: str, packages: list[Package]) -> None:
    output = 'Bitte trage die folgenden Zellen selbst ein:\n'
    for package in packages:
        output = output + f"- {package.coordinate}\n"
    file_name = os.path.join(folder_for_file, "SELBSTEINTRAGEN.txt")
    with open(file_name, "w", encoding="utf-8") as file_out:
        file_out.write(output)
    


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
    xlsx_pattern = r"\.xlsx$"
    beginning_pattern = r"^~\$"
    try:
        for file in os.listdir(settings.not_parsed_excel_file_folder):
            if re.search(xlsx_pattern, file) and not re.search(beginning_pattern, file):
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
    file_name = f"{get_file_name_from_file_path(excel_file).replace('.xlsx', '')}-parsed_at_{now_variable.strftime('%d.%m.%Y_%H-%M-%S')}.xml"
    file_path = os.path.join(settings.xml_output_file_folder, file_name)
    xml_tree = export_manager.get_xml_tree(packages)
    xml_tree.write(file_path, encoding="utf-8", xml_declaration=True)
    return file_name


def has_package_all_needed_informations(package: Package) -> list[bool, str]:
    if not package.recipientName:
        return [False, "Kein Name gefunden"]
    elif not package.address1:
        return [False, "Keine Adresse gefunden"]
    elif not package.country:
        return [False, "Kein Land gefunden"]
    elif not package.postalCode:
        return [False, "Keine PLZ gefunden"]
    elif not package.city:
        return [False, "Keine Stadt gefunden"]
    elif not package.referenceNumbers:
        return [False, "Keine Referenznummer gefunden"]
    else:
        return [True, ""]


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
        coordinate=package.coordinate
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
    new_package.referenceNumbers = package.referenceNumbers
    new_package.packageCount = package.packageCount
    return new_package

def print_adress_error(
    excel_row: int, excel_column: int, number_of_packages: int, packageCount: int
) -> None:
    print_info(f"({str(packageCount)}/{str(number_of_packages)}) Adresse:", 1)
    print_info(
        f"Bitte gib die Adresse in Zelle {int_to_alphabet(excel_column)}{excel_row} manuel ein!",
        2,
        False,
    )


def print_adress_info(
    packages: list[Package], number_of_packages: int, packageCount: int
) -> None:
    package = packages[0]
    print_info(f"({str(packageCount)}/{str(number_of_packages)}) Adresse:", 1)
    print_info(f"Name: {package.recipientName}", 2, False)
    if package.recipientNameAddtional:
        print_info(f"zu H.: {package.recipientNameAddtional}", 2, False)
    print_info(f"Adresse: {package.address1}", 2, False)
    print_info(f"PLZ: {package.postalCode}", 2, False)
    print_info(f"Stadt: {package.city}", 2, False)
    print_info(f"Land: {package.country}", 2, False)
    if package.phoneNumber:
        print_info(f"Telefon: {package.phoneNumber}", 2, False)
    print_info(f"Ref-Nummer: {package.referenceNumbers}", 2, False)
    if len(packages) > 1:
        print_info(
            f"Dieses Paket geht ins Ausland und hat eine Anzahl von {str(len(packages))}. Es wurde zusätzlich {str(len(packages) -1)} mal hinzugefügt",
            2,
            False,
        )
    else:
        print_info(f"Anzahl: {str(package.packageCount)}", 2, False)


def print_adress_info_with_incomplete_address(
    package: Package, missing_value: str, packageCount: int, number_of_packages: int
) -> None:
    print_info(f"({str(packageCount)}/{str(number_of_packages)}) Adresse:", 1)
    print_info(missing_value, 2, False)
    print_info(
        f"Bitte gib die Adresse in Zelle {int_to_alphabet(package.excel_column)}{package.excel_row} manuel ein!",
        2,
        False,
    )
    
    
def get_first_line_from_string(input: str) -> str:
    if "\n" in input:
        split = input.split("\n")
        return split[0]
    else:
        return input
    

def print_result_info(number_of_packages, excel_file_has_a_problem: bool, problem_packages: dict[Package]) -> None:
    if excel_file_has_a_problem: 
        output = f"[{(number_of_packages-len(problem_packages))}/{number_of_packages}] Pakete wurden formatiert. Es gab Fehler bei der/den Zell(en) "
        counter = 0
        for package in problem_packages:
            counter += 1
            output = output + f"{package.coordinate} ({get_first_line_from_string(package.excelReciverString)})"
            if counter == len(problem_packages):
                output += "."
            else:
                output += ","
    else:
        output = f"[{number_of_packages}/{number_of_packages}] Pakete wurden formatiert."
    print_info(output, tab=1)
    
    
def fill_packageName_and_additionalName(package: Package) -> Package:
    if not package.recipientName.strip():
        if package.recipientNameAddtional:
            package.recipientName = package.recipientNameAddtional
        else:
            return package
        
    elif package.country != "DE" and not package.recipientNameAddtional:
        package.recipientNameAddtional = package.recipientName
        package.email = settings.email_placeholder
    
    return package


def main() -> None:
    print_info(f"Das Programm startet.")
    inital_check_on_existing_file_infrastructure()
    print()

    excel_files_to_parse = get_files_to_parse()
    print_excel_file_info(excel_files_to_parse)

    for excel_file in excel_files_to_parse:
        print_info(f"Starte Analyse von '{get_file_name_from_file_path(excel_file)}'")
        excel_file_has_a_problem: list[bool, list] = [False, []]

        try:
            info_and_packages = excel_converter.get_packages_from_excel_file(excel_file)
            excel_type = info_and_packages["type"]
            packages = info_and_packages["packages"]
            excel_file_has_a_problem = info_and_packages["excel_file_has_a_problem"]
            
        except Exception as e:
            print_info(
                f"Fehler beim Auswerten der Excel-Datei '{excel_file}': {e}. Die Datei wird übersprungen!"
            )
            continue

        number_of_packages = len(packages)
        print_info(f"Es wurden {number_of_packages} Pakete gefunden", 1)

        output_packages: list[Package] = []
        packageCount = 0
        for package in packages:
            packageCount += 1
            if excel_type == "old_version":
                try:
                    address_assignment = address_parser.parse_address(
                        package.excelReciverString
                    )
                except Exception as e:
                    excel_file_has_a_problem[0] = True
                    excel_file_has_a_problem[1].append(package)
                    print_adress_error(package.excel_row, package.excel_column, number_of_packages, packageCount)
                    continue

                try:
                    package = address_parser.sort_assignment_to_package(
                        address_assignment, package
                    )
                except Exception as e:
                    excel_file_has_a_problem[0] = True
                    excel_file_has_a_problem[1].append(package)
                    print_adress_error(package.excel_row, package.excel_column, number_of_packages, packageCount)
                    continue
                
            elif excel_type == "new_version":
                package = fill_packageName_and_additionalName(package)

            package_state = has_package_all_needed_informations(package)

            additonal_packages_for_abroads = ckeck_package_on_abroad_and_dublicate(
                package
            )

            if package_state[0]:
                for ready_packages in additonal_packages_for_abroads:
                    output_packages.append(ready_packages)
                print_adress_info(
                    additonal_packages_for_abroads, number_of_packages, packageCount
                )
            else:
                excel_file_has_a_problem[0] = True
                excel_file_has_a_problem[1].append(package)
                print_adress_info_with_incomplete_address(
                    package, package_state[1], packageCount, number_of_packages
                )
            print()

        output_file_name = write_packages_to_xml_file(output_packages, excel_file)
        
        print_result_info(number_of_packages, excel_file_has_a_problem[0], excel_file_has_a_problem[1])

        if excel_file_has_a_problem[0]:
            output_folder = move_file(excel_file)
            if output_folder:
                create_text_file_with_problem_information(output_folder, excel_file_has_a_problem[1])
        else:
            move_file(excel_file)

        print()

    print_info(f"Das Programm ist beendet.")
    print()


main()
