from fuzzywuzzy import fuzz

import settings
from data_classes import Package


def get_highest_country_match(search_part: str) -> list[int, str]:
    highest_ratio = [0, ""]
    for country in settings.european_countrys:
        ratio = fuzz.partial_ratio(search_part, country)
        if ratio > highest_ratio[0]:
            highest_ratio[0] = ratio
            highest_ratio[1] = country
    return highest_ratio


def parse_address(address_string: str) -> dict[str, str]:
    # Split the address string by new lines
    address_parts = address_string.split("\n")
    address_parts_left = address_parts.copy()

    # Initialize the address dictionary with empty values
    address_assignment = {
        "name": None,
        "company": None,
        "street": None,
        "region": None,
        "country": None,
        "tel": None,
    }

    if len(address_parts) >= 7:
        raise Exception(
            f"Fehler beim Adress-Erkennen. Es wurden zu viele Zeilen angegeben."
            f"Maximal 6 Zeilen möglich. Aber es wurden {str(len(address_parts))} Zeilen angegeben\n"
            f"Zeilen: {str(address_parts)}"
        )

    while address_parts_left != []:
        for part in address_parts:
            if not address_assignment["name"] and address_parts.index(part) == 0:
                address_parts_left.remove(address_parts[0])
                address_assignment["name"] = address_parts[0]
                continue

            if not address_assignment["tel"] and settings.phone_pattern.match(part):
                address_parts_left.remove(part)
                address_assignment["tel"] = part
                continue

            if not address_assignment["region"] and settings.region_pattern.match(part):
                address_parts_left.remove(part)
                address_assignment["region"] = part
                continue

            if not address_assignment["street"] and settings.street_pattern.match(part):
                address_parts_left.remove(part)
                address_assignment["street"] = part
                continue

            if not address_assignment["country"] and part in settings.european_countrys:
                address_parts_left.remove(part)
                address_assignment["country"] = part
                continue

            if not address_assignment["country"]:
                highest_country = get_highest_country_match(part)
                if highest_country[0] > 90:
                    address_parts_left.remove(part)
                    address_assignment["country"] = highest_country[1]
                    continue

        if len(address_parts_left) == 1:
            address_assignment["company"] = address_parts_left[0]
            address_parts_left.remove(address_parts_left[0])

        if len(address_parts_left) > 1:
            raise Exception(
                f"Fehler beim Adress-Erkennen. Die Zuordnung war nicht möglich, da zu viele Adressteile nicht erkannt werden können."
                f" Diese Adresse wird übersprungen\nAdresse: {str(address_parts)}"
            )

    return address_assignment


def concatenate_strings_from_second_element(lst: list[str]) -> str:
    result = ""
    for i, item in enumerate(lst):
        if i > 0 and isinstance(item, str):
            result += item + " "
    return result.strip()


def sort_assignment_to_package(address_assignment: dict[str, str], package: Package) -> Package:
    package.recipientName = address_assignment["name"]
    package.address1 = address_assignment["street"]
    package.country = address_assignment["country"]
    package.phoneNumber = address_assignment["tel"]

    package.service = "Standart"
    # TODO: Frage ob das Gewicht bei einem doppeltem Packet verdoppelt werden muss
    package.weight = 10.0

    if address_assignment["company"]:
        package.recipientNameAddtional = address_assignment["company"]
    
    elif address_assignment["country"] != 'Deutschland' and not address_assignment["company"]:
        package.recipientNameAddtional = settings.foreign_country_placeholder

    region = address_assignment["region"]
    region_parts = region.split(",")
    if len(region_parts) > 1:
        if region_parts[1]:
            package.state = region_parts[1]

    region = region_parts[0]
    region_parts = region.split(" ")
    if len(region_parts) == 2:
        package.postalCode = int(region_parts[0])
        package.city = region_parts[1]
    elif len(region_parts) > 2:
        package.postalCode = int(region_parts[0])
        package.city = concatenate_strings_from_second_element(region_parts)
    else:
        raise Exception(f"Es konnte keine PLZ und kein Ort aus '{address_assignment["region"]}' ermittelt werden!")

    return package
