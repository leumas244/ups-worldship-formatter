from fuzzywuzzy import fuzz

import settings
from data_classes import Package


def get_highest_country_match(search_part: str) -> list[int, str]:
    highest_ratio = [0, "", ""]
    for country in settings.european_countrys:
        ratio = fuzz.partial_ratio(search_part, country)
        if ratio > highest_ratio[0]:
            highest_ratio[0] = ratio
            highest_ratio[1] = settings.european_countrys[country]
            highest_ratio[2] = country
    return highest_ratio


def check_on_phonenumber_behind_country(search_part: str, recognized_country: str) -> str:
    try:
        search_part_without_country = search_part.replace(recognized_country, "")
        search_part_without_country = search_part_without_country.strip()
        if settings.phone_pattern.match(search_part_without_country):
            return search_part_without_country
    except:
        if " " in search_part:
            search_part_split = search_part.split(" ")
            for part in search_part_split:
                part = part.strip()
                if settings.phone_pattern.match(part):
                    return search_part_without_country
    return


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
        "country_name_short": None,
        "phonenumber": None,
    }

    if len(address_parts) >= 7:
        raise Exception(
            f"Fehler beim Adress-Erkennen. Es wurden zu viele Zeilen angegeben."
            f"Maximal 6 Zeilen möglich. Aber es wurden {str(len(address_parts))} Zeilen angegeben\n"
            f"Zeilen: {str(address_parts)}"
        )

    while address_parts_left != []:
        for part in address_parts:
            part_strip = part.strip()
            if not address_assignment["name"] and address_parts.index(part) == 0:
                address_parts_left.remove(address_parts[0])
                address_assignment["name"] = address_parts[0]
                continue

            if not address_assignment["phonenumber"] and settings.phone_pattern.match(part_strip):
                address_parts_left.remove(part)
                address_assignment["phonenumber"] = part
                continue

            if not address_assignment["region"] and settings.region_pattern.match(part_strip):
                address_parts_left.remove(part)
                address_assignment["region"] = part
                continue

            if not address_assignment["street"] and settings.street_pattern.match(part_strip):
                address_parts_left.remove(part)
                address_assignment["street"] = part
                continue

            if not address_assignment["country_name_short"] and part_strip in settings.european_countrys:
                address_parts_left.remove(part)
                address_assignment["country_name_short"] = settings.european_countrys[part_strip]
                continue

            if not address_assignment["country_name_short"]:
                highest_country = get_highest_country_match(part_strip)
                if highest_country[0] > 90:
                    phone_number = check_on_phonenumber_behind_country(part_strip, highest_country[2])
                    if phone_number:
                        address_assignment["phonenumber"] = phone_number
                    address_parts_left.remove(part)
                    address_assignment["country_name_short"] = highest_country[1]
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


def get_plz_city_and_region_from_line(line: str) -> dict[str, str]:
    result = {"state": None,
              "postalCode": None,
              "city": None
              }
    region = line.replace(u'\xa0', u' ')
    region_parts = region.split(",")
    if len(region_parts) > 1:
        if region_parts[1]:
            result["state"] = region_parts[1]

    region = region_parts[0]
    region_parts = region.split(" ")
    if len(region_parts) == 2:
        result["postalCode"] = region_parts[0]
        result["city"] = region_parts[1]
        return result
    elif len(region_parts) > 2:
        result["postalCode"] = region_parts[0]
        result["city"] = concatenate_strings_from_second_element(region_parts)
        return result
    else:
        raise Exception(f"Es konnte keine PLZ und kein Ort aus '{line}' ermittelt werden!")


def sort_assignment_to_package(address_assignment: dict[str, str], package: Package) -> Package:
    package.recipientName = address_assignment["name"]
    package.address1 = address_assignment["street"]
    package.country = address_assignment["country_name_short"]
    package.phoneNumber = address_assignment["phonenumber"]

    package.service = "Standart"
    package.weight = 10.0 * package.packageCount

    if address_assignment["company"]:
        package.recipientNameAddtional = address_assignment["company"]
    
    elif address_assignment["country_name_short"] != "DE" and not address_assignment["company"]:
        package.recipientNameAddtional = package.recipientName

    region_info = get_plz_city_and_region_from_line(address_assignment["region"])
    package.state = region_info["state"]
    package.postalCode = region_info["postalCode"]
    package.city = region_info["city"]

    return package
