import xml.etree.ElementTree as ET
import os
import shutil
from datetime import datetime

from data_classes import Package

def get_xml_tree(packages: list[Package]) -> ET.ElementTree:
    OpenShipments = ET.Element("OpenShipments")
    OpenShipments.set("xmlns", "x-schema:OpenShipments.xdr")
    
    for package in packages:
        OpenShipment = ET.SubElement(OpenShipments, "OpenShipment")
        OpenShipment.set("ShipmentOption", "")
        OpenShipment.set("ProcessStatus", "")
        
        ShipTo = ET.SubElement(OpenShipment, "ShipTo")
        
        CompanyOrName = ET.SubElement(ShipTo, "CompanyOrName")
        CompanyOrName.text = package.recipientName
        
        if package.recipientNameAddtional:
            Attention = ET.SubElement(ShipTo, "Attention")
            Attention.text = package.recipientNameAddtional
            
        Address1 = ET.SubElement(ShipTo, "Address1")
        Address1.text = package.address1
        
        CityOrTown = ET.SubElement(ShipTo, "CityOrTown")
        CityOrTown.text = package.city
        
        if package.country == "Deutschland":
            CountryTerritory = ET.SubElement(ShipTo, "CountryTerritory")
            CountryTerritory.text = "DE"
        
        PostalCode = ET.SubElement(ShipTo, "PostalCode")
        PostalCode.text = str(package.postalCode)
        
        StateProvinceCounty = ET.SubElement(ShipTo, "StateProvinceCounty")
        StateProvinceCounty.text = package.state
        
        Telephone = ET.SubElement(ShipTo, "Telephone")
        Telephone.text = package.phoneNumber
        
        ShipFrom = ET.SubElement(OpenShipment, "ShipFrom")
        CompanyOrName_from = ET.SubElement(ShipFrom, "CompanyOrName")
        CompanyOrName_from.text = "Wildstage GmbH"
        Attention_from = ET.SubElement(ShipFrom, "Attention")
        Attention_from.text = "Kai Funk"
        Address1_from = ET.SubElement(ShipFrom, "Address1")
        Address1_from.text = "Alleestr. 15-19"
        CountryTerritory_from = ET.SubElement(ShipFrom, "CountryTerritory")
        CountryTerritory_from.text = "DE"
        PostalCode_from = ET.SubElement(ShipFrom, "PostalCode")
        PostalCode_from.text = "33818"
        CityOrTown_from = ET.SubElement(ShipFrom, "CityOrTown")
        CityOrTown_from.text = "Leopoldshöhe"
        
        ShipmentInformation = ET.SubElement(OpenShipment, "ShipmentInformation")
        ServiceType = ET.SubElement(ShipmentInformation, "ServiceType")
        ServiceType.text = "Standard"
        NumberOfPackages = ET.SubElement(ShipmentInformation, "NumberOfPackages")
        NumberOfPackages.text = str(package.packageCount)
        ShipmentActualWeight = ET.SubElement(ShipmentInformation, "ShipmentActualWeight")
        ShipmentActualWeight.text = str(package.weight).replace(".", ",")
        
        Package = ET.SubElement(OpenShipment, "Package")
        PackageType = ET.SubElement(Package, "PackageType")
        PackageType.text = "Package"
        Weight = ET.SubElement(Package, "Weight")
        Weight.text = str(package.weight).replace(".", ",")
        Reference1 = ET.SubElement(Package, "Reference1")
        Reference1.text = package.referenceNumber
        Length = ET.SubElement(Package, "Length")
        Length.text = "0"
        Width = ET.SubElement(Package, "Width")
        Width.text = "0"
        Height = ET.SubElement(Package, "Height")
        Height.text = "0"
        
    tree = ET.ElementTree(OpenShipments)
    
    return tree