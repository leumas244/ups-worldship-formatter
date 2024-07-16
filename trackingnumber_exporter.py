import xml.etree.ElementTree as ElementTree
import os
import re

import settings

def get_file_name_from_file_path(path_to_file: str) -> str:
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
        
def get_matching_excel_and_out_files() -> list[tuple[str]]:
    pass
        
def get_xml_tree(file_path: str) -> ElementTree:
    tree = ElementTree.parse(file_path)
    return tree.getroot()
        
def start_routine() -> None:
    proccesed_xml_files = get_all_processed_xml_files()
    
    excel_and_out_files = get_matching_excel_and_out_files()
    
    for proccesed_xml_file in proccesed_xml_files:
        xml_root = get_xml_tree(proccesed_xml_file)
        
        openShipments = xml_root.findall("{x-schema:OpenShipments.xdr}OpenShipment")
        
        for openShipment in openShipments:
            process_status = openShipment.get('ProcessStatus')
            if process_status == "Processed":
                
                pass