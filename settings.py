import re

## Ordner
# Ordner für Excel-Datei, die geparst werden sollen.
not_parsed_excel_file_folder: str = "data/noch_NICHT_analysierte_Excel_Dateien/"
# Ordner in denen die fertig geparsten Excel-Datei verschoben werden.
parsed_excel_file_folder: str = "data/analysierte_Excel_Dateien/"
# Ordner für Excel-Datei, bei denen es ein Problem gab.
parsed_excel_file_with_problems_folder: str = "data/analysierte_Excel_Dateien_mit_PROBLEMEN/"
# Ordner für csv Datei Ausgabe.
csv_output_file_folder: str = "data/csv_Ausgabe_Dateien_fuer_UPS-WorldShip/"

all_needed_folders: list[str] = [
    not_parsed_excel_file_folder,
    parsed_excel_file_folder,
    parsed_excel_file_with_problems_folder,
    csv_output_file_folder
]

## Anpassbare Variablen/Einstellungen
# reguläre Ausdrücker
street_pattern = re.compile(r'^[^\d /-]{1}.+ \d+[a-zA-Z]?$')
region_pattern = re.compile(r'^\d{4,5} [^\d/-]{1}.+$')
phone_pattern = re.compile(r"^\+?\d+([ ]?[/]?[ ]?\d+)*$")

# Platzhalter für zusätzlichen Namen, der nötig ist ausserhalb Deutschlands
foreign_country_placeholder: str = "Auslandsplatzhalter"

# Erkennbare Länder
european_countrys: list[str] = [
    "Deutschland",
    "Österreich",
    "Schweiz",
    "Luxemburg",
    "Luxembourg",
    "Niederlande",
    "Nederland",
    "Frankreich",
    "France",
    "Belgien",
]
