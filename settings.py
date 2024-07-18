import re

## Ordner
# Ordner für Excel-Datei, die geparst werden sollen.
not_parsed_excel_file_folder: str = "data/noch_NICHT_analysierte_Excel_Dateien/"
# Ordner in denen die fertig geparsten Excel-Datei verschoben werden.
parsed_excel_file_folder: str = "data/analysierte_Excel_Dateien/"
# Ordner für Excel-Datei, bei denen es ein Problem gab.
ups_protocol_folder: str = (
    "data/UPS-Protokolle/"
)
# Ordner für csv Datei Ausgabe.
xml_output_file_folder: str = "data/xml_Ausgabe_Dateien_fuer_UPS-WorldShip/"

all_needed_folders: list[str] = [
    not_parsed_excel_file_folder,
    parsed_excel_file_folder,
    ups_protocol_folder,
    xml_output_file_folder,
]

## Anpassbare Variablen/Einstellungen
# reguläre Ausdrücker
street_pattern = re.compile(r"^[^\d /-]{1}.+ \d+([a-zA-Z]?(/?\d+)*)$")
region_pattern = re.compile(r"^\d{4,5} [^\d/-]{1}.+$")
phone_pattern = re.compile(r"^\+?\d+([ ]?[/-]?[ ]?\d+)*$")

# Erkennbare Länder
european_countrys: dict[str, str] = {
    "Deutschland": "DE",
    "Germany": "DE",
    "Dänemark": "DK",
    "Denmark": "DK",
    "Danmark": "DK",
    "Polen": "PL",
    "Poland": "PL",
    "Polska": "PL",
    "Tschechien": "CZ",
    "Tschechische Republik": "CZ",
    "Czech Republic": "CZ",
    "Česká republika": "CZ",
    "Österreich": "AT",
    "Austria": "AT",
    "Schweiz": "CH",
    "Switzerland": "CH",
    "Suisse": "CH",
    "Frankreich": "FR",
    "France": "FR",
    "Luxemburg": "LU",
    "Luxembourg": "LU",
    "Lëtzebuerg": "LU",
    "Belgien": "BE",
    "Belgium": "BE",
    "Belgique": "BE",
    "Niederlande": "NL",
    "Netherlands": "NL",
    "Nederland": "NL",
}
