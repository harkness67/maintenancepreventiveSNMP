from pysnmp.hlapi import *
from datetime import datetime
import logging
logging.basicConfig(filename='printer_retrieval.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


# Dictionnaire des OIDs pour chaque constructeur
OID_MAP = {
    "HP": {
        "name": "1.3.6.1.2.1.1.5.0",
        "serial_number": "1.3.6.1.2.1.43.5.1.1.17.1",
        "page_counter": "1.3.6.1.2.1.43.10.2.1.4.1.1"
    },
    "Samsung": {
        "name": ".1.3.6.1.4.1.236.11.5.11.55.5.1.6",
        "serial_number": "1.3.6.1.4.1.236.11.5.11.55.5.1.4",
        "page_counter": "1.3.6.1.4.1.236.11.5.11.55.20.2.1.9"
    },
    "Lexmark": {
        "name": "1.3.6.1.4.1.641",
        "serial_number": "1.3.6.1.4.1.641.6.2.3.1.5",
        "page_counter": "1.3.6.1.4.1.641.6.4.4.1.1.11"
    }
}

def get_printer_name(community, ip):
    # Essaie de récupérer le nom de l'imprimante pour chaque fabricant défini dans OID_MAP
    for manufacturer, oids in OID_MAP.items():
        oid = oids["name"]
        try:
            errorIndication, errorStatus, errorIndex, varBinds = next(
                getCmd(SnmpEngine(),
                       CommunityData(community, mpModel=1),
                       UdpTransportTarget((ip, 161)),
                       ContextData(),
                       ObjectType(ObjectIdentity(oid))
                       ))

            if not errorIndication and varBinds:
                return varBinds[0][1].prettyPrint(), manufacturer
        except Exception as e:
            print(f"Erreur lors de la récupération du nom de l'imprimante pour l'adresse IP {ip}: {str(e)}")
            continue  # Passez à l'OID suivant

    # Si aucune réponse n'est obtenue pour tous les OIDs, renvoie None pour les deux
    return None, None

def detect_manufacturer(community, ip):
    _, manufacturer = get_printer_name(community, ip)
    return manufacturer

def get_printer_info(ip_address, community='public'):
    manufacturer = detect_manufacturer(community, ip_address)
    
    if not manufacturer:
        logging.warning(f"Impossible de détecter le fabricant pour l'adresse IP {ip_address}.")
        return None
        print(f"Impossible de détecter le fabricant pour l'adresse IP {ip_address}.")
        return None
    
    name_oid = OID_MAP[manufacturer]["name"]
    serial_oid = OID_MAP[manufacturer]["serial_number"]
    page_counter_oid = OID_MAP[manufacturer]["page_counter"]
    
    try:
        # Ouvrir une session SNMP avec l'imprimante pour obtenir les informations
        errorIndication, errorStatus, errorIndex, varBinds = next(
            getCmd(SnmpEngine(),
                   CommunityData(community, mpModel=1),
                   UdpTransportTarget((ip_address, 161), timeout=2.0),
                   ContextData(),
                   ObjectType(ObjectIdentity(name_oid)),
                   ObjectType(ObjectIdentity(serial_oid)),
                   ObjectType(ObjectIdentity(page_counter_oid))
                   ))

        if errorIndication:
            logging.error(f"Erreur SNMP lors de la requête à l'adresse {ip_address}: {errorIndication}")
            print(f"Erreur SNMP lors de la requête à l'adresse {ip_address}: {errorIndication}")
            return None
        
        sys_name = varBinds[0][1].prettyPrint()
        serial_number = varBinds[1][1].prettyPrint()
        page_counter = varBinds[2][1].prettyPrint()
        collect_date = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        
        return sys_name, serial_number, page_counter, collect_date

    except Exception as e:
        logging.error(f"Erreur lors de la récupération des informations pour l'adresse IP {ip_address}: {str(e)}")
        print(f"Erreur lors de la récupération des informations pour l'adresse IP {ip_address}: {str(e)}")
        return None