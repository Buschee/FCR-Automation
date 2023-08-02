import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import requests
from requests.auth import HTTPBasicAuth
from getpass import getpass
import json
import urllib3
import datetime
import socket
import struct
import re
import sys
import time

#Global parameter
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
#test mgmt
#mgmt_ip = "192.168.1.1"
mgmt_ip = "213.33.120.99"


def print_banner():
   print("""  ______ _____ _____                    _                        _   _             
 |  ____/ ____|  __ \        /\        | |                      | | (_)            
 | |__ | |    | |__) |_____ /  \  _   _| |_ ___  _ __ ___   __ _| |_ _  ___  _ __  
 |  __|| |    |  _  /______/ /\ \| | | | __/ _ \| '_ ` _ \ / _` | __| |/ _ \| '_ \ 
 | |   | |____| | \ \     / ____ \ |_| | || (_) | | | | | | (_| | |_| | (_) | | | |
 |_|    \_____|_|  \_\   /_/    \_\__,_|\__\___/|_| |_| |_|\__,_|\__|_|\___/|_| |_| by DHaslauer
 -----------------------------------------------------------------------------------------------
                                                                            """)
   

def select_excel_file():
    """
    Öffnet ein Dateiauswahlfenster, um eine Excel-Datei auszuwählen
    und gibt den Pfad zur ausgewählten Datei zurück.
    """
    root = Tk()
    root.withdraw()
    excel_file = askopenfilename(filetypes=[("Excel-Dateien", ".xlsx .xls")])
    return excel_file


def extract_data(excel_file):
    """
    Extrahiert die Informationen ab der "ID"-Zeile aus einer Excel-Datei,
    die durch den Pfad 'excel_file' definiert ist, und gibt ein Pandas DataFrame zurück.
    """
    df = pd.read_excel(excel_file, header=None, skiprows=6)

    # Finde die Zeile, die mit dem String "ID" beginnt
    id_row = df[df.apply(lambda row: row.astype(str).str.contains('ID', case=False).any(), axis=1)]

    # Überprüfe, ob id_row definiert wurde, bevor du die Daten extrahierst
    if not id_row.empty:
        # Extrahiere die Informationen ab der "ID"-Zeile
        start_row = id_row.index[0]
        data = df.iloc[start_row:]
        return data
    else:
        print("[!] ID-row not found")
        return None

def connect_checkpoint_api():
    """
    Stellt eine Verbindung zur Check Point API her und gibt die
    Authentifizierungs-Header zurück.
    """
    # Abfrage des API-Passworts vom Benutzer
    api_password = getpass("[*] Please enter the password for the api user: ")

    # Verbindung zur Check Point API herstellen
    url = f"https://{mgmt_ip}/web_api/login"
    payload = {'user': 'fcr-user', 'password': api_password}
    response = requests.post(url, json=payload, verify=False)
    if response.status_code == 200:
        # Authentifizierungs-Header zurückgeben
        #auth_header = {"Authorization": "Bearer " + response.json()['sid']}
        auth_header = {
            "Content-Type": "application/json",
            "X-chkp-sid": response.json()['sid']
        }
        print("[+] Successfully connected to Checkpoint API")
        return auth_header
    else:
        print("[!] Error while authentication")
        print("[!] Error: ", response.json()['message'])
        quit()
        return None
    

def get_firewall_objects(data):
    """
    Extrahiert die Firewall-Objekte aus den übergebenen Daten.
    Diese werden in den Spalten 6 und 9 gesucht und dürfen nicht "NaN" sein.
    Gibt eine Liste der gefundenen Firewall-Objekte zurück.
    """
    objects = []
    for index, row in data.iterrows():
        object1 = str(row[6])
        object2 = str(row[9])
        if object1 != 'nan':
            objects += object1.split('\n')
            object1 = str(object1.replace("\n", ""))
        if object2 != 'nan':
            objects += object2.split('\n')
            object2 = str(object2.replace("\n", ""))
    return objects


def get_sources(data):
    objects = []
    for index, row in data.iterrows():
        object1 = str(row[6])
        if object1 != 'nan':
            if object1 != '':
                objects += object1.split('\n')
                object1 = str(object1.replace("\n", ""))
                object1 = str(object1.replace(" ", ""))
                object1 = str(object1.replace("NaN", ""))
                object1 = str(object1.replace("nan", ""))
                object1 = str(object1.replace("Add", ""))
                object1 = str(object1.replace("add", ""))
                object1 = str(object1.replace("ADD", ""))
                object1 = str(object1.replace("Add:", ""))
                object1 = str(object1.replace("add:", ""))
                object1 = str(object1.replace("ADD:", ""))
                object1 = str(object1.replace("Del", ""))
                object1 = str(object1.replace("del", ""))
                object1 = str(object1.replace("DEL", ""))
                object1 = str(object1.replace("Del:", ""))
                object1 = str(object1.replace("del:", ""))
                object1 = str(object1.replace("DEL:", ""))
                object1 = str(object1.replace("delete", ""))
                object1 = str(object1.replace("delete:", ""))
                object1 = str(object1.replace("Delete", ""))
                object1 = str(object1.replace("Delete:", ""))
                object1 = str(object1.replace("DELETE", ""))
                object1 = str(object1.replace("DELETE:", ""))
                
    return objects

def get_destinations(data):
    objects = []
    for index, row in data.iterrows():
        object2 = str(row[9])
        if object2 != 'nan':
            if object2 != '':
                objects += object2.split('\n')
                object2 = str(object2.replace("\n", ""))
                object2 = str(object2.replace(" ", ""))
                object2 = str(object2.replace("NaN", ""))
                object2 = str(object2.replace("nan", ""))
                object2 = str(object2.replace("Add", ""))
                object2 = str(object2.replace("add", ""))
                object2 = str(object2.replace("ADD", ""))
                object2 = str(object2.replace("Add:", ""))
                object2 = str(object2.replace("add:", ""))
                object2 = str(object2.replace("ADD:", ""))
                object2 = str(object2.replace("Del", ""))
                object2 = str(object2.replace("del", ""))
                object2 = str(object2.replace("DEL", ""))
                object2 = str(object2.replace("Del:", ""))
                object2 = str(object2.replace("del:", ""))
                object2 = str(object2.replace("DEL:", ""))
                object2 = str(object2.replace("delete", ""))
                object2 = str(object2.replace("delete:", ""))
                object2 = str(object2.replace("Delete", ""))
                object2 = str(object2.replace("Delete:", ""))
                object2 = str(object2.replace("DELETE", ""))
                object2 = str(object2.replace("DELETE:", ""))
    return objects


def get_service_objects(data): 
    """
    Extrahiert die Service-Objekte aus den übergebenen Daten.
    Diese werden in der Spalte 12 gesucht und dürfen nicht "NaN" sein.
    Gibt eine Liste der gefundenen Service-Objekte zurück.
    """
    objects = []
    for index, row in data.iterrows():
        object1 = str(row[12])
        if object1 != 'nan':
            objects += object1.split('\n')
            object1 = str(object1.replace("\n", ""))
            object1 = str(object1.replace(" ", ""))
            object1 = str(object1.replace("Add", ""))
            object1 = str(object1.replace("add", ""))
            object1 = str(object1.replace("ADD", ""))
            object1 = str(object1.replace("Add:", ""))
            object1 = str(object1.replace("add:", ""))
            object1 = str(object1.replace("ADD:", ""))
            object1 = str(object1.replace("Del", ""))
            object1 = str(object1.replace("del", ""))
            object1 = str(object1.replace("DEL", ""))
            object1 = str(object1.replace("Del:", ""))
            object1 = str(object1.replace("del:", ""))
            object1 = str(object1.replace("DEL:", ""))
            object1 = str(object1.replace("delete", ""))
            object1 = str(object1.replace("delete:", ""))
            object1 = str(object1.replace("Delete", ""))
            object1 = str(object1.replace("Delete:", ""))
            object1 = str(object1.replace("DELETE", ""))
            object1 = str(object1.replace("DELETE:", ""))
    return objects


def get_installation_gw(data):
    objects = []
    for index, row in data.iterrows():
        object1 = str(row[1])
        if object1 != 'nan':
            objects += object1.split('\n')
            object1 = str(object1.replace("\n", ""))
            object1 = str(object1.replace(" ", ""))
            object1 = str(object1.replace("Add", ""))
            object1 = str(object1.replace("add", ""))
            object1 = str(object1.replace("ADD", ""))
            object1 = str(object1.replace("Add:", ""))
            object1 = str(object1.replace("add:", ""))
            object1 = str(object1.replace("ADD:", ""))
            object1 = str(object1.replace("Del", ""))
            object1 = str(object1.replace("del", ""))
            object1 = str(object1.replace("DEL", ""))
            object1 = str(object1.replace("Del:", ""))
            object1 = str(object1.replace("del:", ""))
            object1 = str(object1.replace("DEL:", ""))
            object1 = str(object1.replace("delete", ""))
            object1 = str(object1.replace("delete:", ""))
            object1 = str(object1.replace("Delete", ""))
            object1 = str(object1.replace("Delete:", ""))
            object1 = str(object1.replace("DELETE", ""))
            object1 = str(object1.replace("DELETE:", ""))
    return objects


def check_objects_exist(auth_header, object_names):
    """
    Überprüft, ob die übergebenen Firewall-Objektnamen bereits in Checkpoint vorhanden sind.
    Gibt eine Liste der Namen zurück, die noch nicht existieren.
    """
    #existing_objects = []
    not_existing_objects = []
    for object_name in object_names:
        object_name = str(object_name.replace("\n", ""))
        # Überprüfung für Hosts
        if object_name.startswith("h"):
            url = f"https://{mgmt_ip}/web_api/show-host"
            payload = {
                "name": object_name
            }
            data = json.dumps(payload)
            response = requests.post(url, data=data, headers=auth_header, verify=False)
            response_json = json.loads(response.text)
            if response.status_code != 200:
                print(f"[!] {response_json}")
            message = f"Requested object [{object_name}] not found"
            if str(message) in str(response_json):
                # Das Objekt existiert noch nicht
                not_existing_objects.append(object_name)

        # Überprüfung für Networks
        elif object_name.startswith("n"):
            url = f"https://{mgmt_ip}/web_api/show-network"
            payload = {
                "name": object_name
            }
            data = json.dumps(payload)
            response = requests.post(url, data=data, headers=auth_header, verify=False)
            response_json = json.loads(response.text)
            if response.status_code != 200:
                print(f"[!] {response_json}")
            message = f"Requested object [{object_name}] not found"
            if str(message) in str(response_json):
                # Das Objekt existiert noch nicht
                not_existing_objects.append(object_name)

        # Überprüfung für Gruppen
        elif object_name.startswith("g"):
            url = f"https://{mgmt_ip}/web_api/show-group"
            payload = {
                "name": object_name
            }
            data = json.dumps(payload)
            response = requests.post(url, data=data, headers=auth_header, verify=False)
            response_json = json.loads(response.text)
            if response.status_code != 200:
                print(f"[!] {response_json}")
            message = f"Requested object [{object_name}] not found"
            if str(message) in str(response_json):
                # Das Objekt existiert noch nicht
                not_existing_objects.append(object_name)

        # Überprüfung für Active Directory Groups
        elif object_name.startswith("adg"):
            url = f"https://{mgmt_ip}/web_api/show-access-role"
            payload = {
                "name": object_name
            }
            data = json.dumps(payload)
            response = requests.post(url, data=data, headers=auth_header, verify=False)
            response_json = json.loads(response.text)
            if response.status_code != 200:
                print(f"[!] {response_json}")
            message = f"Requested object [{object_name}] not found"
            if str(message) in str(response_json):
                # Das Objekt existiert noch nicht
                not_existing_objects.append(object_name)

        # Überprüfung VPN Access Roles
        elif object_name.startswith("VPN"):
            url = f"https://{mgmt_ip}/web_api/show-access-role"
            payload = {
                "name": object_name
            }
            data = json.dumps(payload)
            response = requests.post(url, data=data, headers=auth_header, verify=False)
            response_json = json.loads(response.text)
            if response.status_code != 200:
                print(f"[!] {response_json}")
            message = f"Requested object [{object_name}] not found"
            if str(message) in str(response_json):
                # Das Objekt existiert noch nicht
                not_existing_objects.append(object_name)

        # Unbekannter Objekt Typ
        else:
            print("[!] Unknown object type...")

    if not_existing_objects:
        print(f"[*] New objects will be created: {not_existing_objects}")
    else:
        print("[*] All objects have already been created, no additional objects will be created.")
    # Rückgabe der Liste der nicht existierenden Objekte
    return not_existing_objects


# Überprüfung ob Services bereits existieren
def check_services_exist(auth_header, service_objects):
    print(service_objects)
    not_existing_services = []
    for service_name in service_objects:
        service_name = str(service_name.replace("\n", ""))
        service_name = str(service_name.replace("Add", ""))
        service_name = str(service_name.replace("add", ""))
        service_name = str(service_name.replace("ADD", ""))
        service_name = str(service_name.replace("Add:", ""))
        service_name = str(service_name.replace("add:", ""))
        service_name = str(service_name.replace("ADD:", ""))
        service_name = str(service_name.replace("Del", ""))
        service_name = str(service_name.replace("del", ""))
        service_name = str(service_name.replace("DEL", ""))
        service_name = str(service_name.replace("Del:", ""))
        service_name = str(service_name.replace("del:", ""))
        service_name = str(service_name.replace("DEL:", ""))
        service_name = str(service_name.replace("delete", ""))
        service_name = str(service_name.replace("delete:", ""))
        service_name = str(service_name.replace("Delete", ""))
        service_name = str(service_name.replace("Delete:", ""))
        service_name = str(service_name.replace("DELETE", ""))
        service_name = str(service_name.replace("DELETE:", ""))
        service_name = str(service_name.replace(":", ""))
        if "tcp" in service_name:
            # Search for tcp service
            url = f"https://{mgmt_ip}/web_api/show-service-tcp"
            payload = {
                "name": service_name
            }
            data = json.dumps(payload)
            response = requests.post(url, data=data, headers=auth_header, verify=False)
            response_json = json.loads(response.text)
            if response.status_code != 200:
                print(f"[!] {response_json}")
            message = f"Requested object [{service_name}] not found"
            if str(message) in str(response_json):
                # Der Service existiert noch nicht
                not_existing_services.append(service_name)

        elif "udp" in service_name:
            # Search or udp service
            url = f"https://{mgmt_ip}/web_api/show-service-udp"
            payload = {
                "name": service_name
            }
            data = json.dumps(payload)
            response = requests.post(url, data=data, headers=auth_header, verify=False)
            response_json = json.loads(response.text)
            if response.status_code != 200:
                print(f"[!] {response_json}")
            message = f"Requested object [{service_name}] not found"
            if str(message) in str(response_json):
                # Der Service existiert noch nicht
                not_existing_services.append(service_name)

        elif service_name == "":
            print("")

        # Check for tcp service
        else:
            # Search for tcp service
            url = f"https://{mgmt_ip}/web_api/show-service-tcp"
            payload = {
                "name": service_name
            }
            data = json.dumps(payload)
            response = requests.post(url, data=data, headers=auth_header, verify=False)
            response_json = json.loads(response.text)
            if response.status_code != 200:
                print(f"[!] {response_json}")
            message = f"Requested object [{service_name}] not found"
            if str(message) in str(response_json):
                # Der Service existiert noch nicht
                not_existing_services.append(service_name)

    if not_existing_services:
        print(f"[*] New services will be created: {not_existing_services}")
    else:
        print("[*] All services have already been created, no additional services will be created.")
    # Rückgabe der Liste der nicht existierenden Services
    return not_existing_services


def create_object(auth_header, name):
    # line break aus Namen entfernen
    name = str(name.replace("\n", ""))
    # create host object
    if name.startswith("h"): 
        print(f"[*] Creating host object {name} with the following parameters")
        # Host Objekt in folgendem Format: h_{Wichtigkeit}_{IP}_{Name}
        split_string = name.split("_")
        ip_value = split_string[2]
        #print(f"The IP of the object is {ip_value}")
        payload = {
            "name": name,
            "ipv4-address": ip_value,
            "color": "black"
        }
        url = f"https://{mgmt_ip}/web_api/add-host"
        response = requests.post(url, data=json.dumps(payload), headers=auth_header, verify=False)
        response_json = json.loads(response.text)
        if response.status_code != 200:
            print(f"[!] {response_json}")
            discard_changes(auth_header)

    # create network object
    elif name.startswith("n"):
        print(f"[*] Creating network object {name}...")
        # Network Objekt in forlgendem Format: n_{Wichtigkeit}_{IP}_{SM}_{Name}
        split_string = name.split("_")
        ip_value = split_string[2]
        # calculate netmask
        subnetmask_value = split_string[3]
        host_bits = 32 - int(subnetmask_value)
        netmask = socket.inet_ntoa(struct.pack('!I', (1 << 32) - (1 << host_bits)))
        # create object
        #print(f"The IP of the network is {ip_value} | {netmask}")
        payload = {
            "name": name,
            "subnet": ip_value,
            "subnet-mask": netmask,
            "color": "black" 
        }
        url = f"https://{mgmt_ip}/web_api/add-network"
        response = requests.post(url, data=json.dumps(payload), headers=auth_header, verify=False)
        response_json = json.loads(response.text)
        if response.status_code != 200:
            print(f"[!] {response_json}")
            discard_changes(auth_header)

    # create group object
    elif name.startswith("g"):
        print(f"[*] Creating group object {name}...")
        # Network Objekt in forlgendem Format: n_{Wichtigkeit}_{Name}
        payload = {
            "name": name
        }
        url = f"https://{mgmt_ip}/web_api/add-group"
        response = requests.post(url, data=json.dumps(payload), headers=auth_header, verify=False)
        response_json = json.loads(response.text)
        if response.status_code != 200:
            print(f"[!] {response_json}")
            discard_changes(auth_header)
        print("[!] Members of the new group object have to be added manually")

    elif name.startswith("adg"):
        print(f"[*] Creating network object {name}...")
        # Access Role Objekt in folgendem Format: adg_{x}_{x}_{x}_{Gruppenname}
        split_string = name.split("_")
        group_name = "_".join(split_string[4:])
        payload = {
            "name": name,
            "color": "black",
            "users": {
                "source": "LDAP Groups",
                "selection": group_name,
                "base-dn": group_name
            }
        }
        url = f"https://{mgmt_ip}/web_api/add-access-role"
        response = requests.post(url, data=json.dumps(payload), headers=auth_header, verify=False)
        response_json = json.loads(response.text)
        if response.status_code != 200:
            print(f"[!] {response_json}")
            discard_changes(auth_header)

    elif name.startswith("VPN"):
        print(f"[*] Creating network object {name}...")
        # Access Role Objekt in folgendem Format: VPN_{Gruppenname}
        split_string = name.split("_")
        group_name = "_".join(split_string[1:])
        payload = {
            "name": name,
            "color": "black",
            "users": {
                "source": "LDAP Groups",
                "selection": group_name,
                "base-dn": group_name
            }
        }
        url = f"https://{mgmt_ip}/web_api/add-access-role"
        response = requests.post(url, data=json.dumps(payload), headers=auth_header, verify=False)
        response_json = json.loads(response.text)
        if response.status_code != 200:
            print(f"[!] {response_json}")
            discard_changes(auth_header)


    #Unknown object type
    else:
        print(f"[!] Unknown object type: {name}")


def create_service(auth_header, name):
    # line break aus Namen entfernen
    name = str(name.replace("\n", ""))
    # create host object
    if "tcp" in name:
        # creating tcp service
        # Format: tcp_{Portnummer}
        split_string = name.split("_")
        service_number = split_string[1]
        payload = {
            "name": name,
            "port": str(service_number)
        }
        url = f"https://{mgmt_ip}/web_api/add-service-tcp"
        response = requests.post(url, data=json.dumps(payload), headers=auth_header, verify=False)
        response_json = json.loads(response.text)
        if response.status_code != 200:
            print(f"[!] {response_json}")
            discard_changes(auth_header)
    
    elif "udp" in name:
        # creating udp service
        # Format: udp_{Portnummer}
        split_string = name.split("_")
        service_number = split_string[1]
        payload = {
            "name": name,
            "port": str(service_number)
        }
        url = f"https://{mgmt_ip}/web_api/add-service-udp"
        response = requests.post(url, data=json.dumps(payload), headers=auth_header, verify=False)
        response_json = json.loads(response.text)
        if response.status_code != 200:
            print(f"[!] {response_json}")
            discard_changes(auth_header)
    
    else:
        # creating tcp service
        # FOrmat: {Name}
        payload = {
            "name": name
        }
        url = f"https://{mgmt_ip}/web_api/add-service-tcp"
        response = requests.post(url, data=json.dumps(payload), headers=auth_header, verify=False)
        response_json = json.loads(response.text)
        if response.status_code != 200:
            print(f"[!] {response_json}")
            discard_changes(auth_header)


def get_rule_position(section_name, auth_header):
    print(section_name)
    if "(" in section_name:
        section_search = re.sub('\(.*?\)', '', section_name)
        section_search = section_search.rsplit(" ", 1)[0].rstrip()
        section_name = section_search
        print(f"[*] Section name: {section_name}")
    else:
        print(f"[*] Section name: {section_name}")
    layer = "Network"
    payload = {
        "name": section_name,
        "layer": layer
    }
    url = f"https://{mgmt_ip}/web_api/show-access-section"
    response = requests.post(url, data=json.dumps(payload), headers=auth_header, verify=False)
    response_json = json.loads(response.text)
    print(f" [*] {response_json}")
    return response_json["uid"]


def get_section_name(section_name):
    print(f"[*] Searching for section: {section_name}")
    section_search = re.sub('\(.*?\)', '', section_name)
    section_search = section_search.rsplit(" ", 1)[0].rstrip()
    return section_search

def get_rule_name(section_name, section_id, auth_header):
    section_search = re.sub('\(.*?\)', '', section_name)
    section_search = section_search.rsplit(" ", 1)[0].rstrip()


def get_rule_id(auth_header, rule_name):
    print(f"[*] Searching for rule: {rule_name}")
    layer = "Network"
    payload = {
        "name": rule_name,
        "layer": layer
    }
    url = f"https://{mgmt_ip}/web_api/show-access-rule"
    response = requests.post(url, data=json.dumps(payload), headers=auth_header, verify=False)
    response_json = json.loads(response.text)
    print(f"[*] {response_json}")
    return response_json["uid"]

def get_description(auth_header, rule_uid):
    layer = "Network"
    payload = {
        "uid": rule_uid,
        "layer": layer
    }
    url = f"https://{mgmt_ip}/web_api/show-access-rule"
    response = requests.post(url, data=json.dumps(payload), headers=auth_header, verify=False)
    response_json = json.loads(response.text)
    print(f"[*] {response_json}")
    return response_json["comments"]


def create_change_rules(auth_header, data, session_name):
    #print(data)
    for index, row in data.iterrows():
        if str(row[0]).startswith('ID'):
            section_name = str(row[0])
            print(f"[*] Section: {section_name}")
        else:
            # read add, change, delete cell
            rule_type = str(row[5])
            rule_type = str(rule_type.replace("\n", ""))
            rule_type = str(rule_type.replace(" ", ""))
            rule_type = str(rule_type.replace("NaN", ""))
            rule_type = str(rule_type.replace("nan", ""))

            if rule_type == "ADD" or rule_type == "add" or rule_type == "" or rule_type == "Add":
                if section_name:
                    #New rule will be created
                    #print(rule_type)
                    #print("Daten aus Excel auslesen und Rule anlegen.")
                    # Daten aus Zeile auslesen
                    section_id = get_rule_position(section_name, auth_header)
                    #get installation targets
                    install_on = []
                    install_on1 = str(row[1])
                    if install_on1 != 'nan':
                        install_on += install_on1.split('\n')
                        install_on1 = str(install_on1.replace("\n", ""))
                        install_on1 = str(install_on1.replace(" ", ""))
                    install_on = list(filter(None, install_on))
                    #print(install_on)

                    #get cell sources
                    sources = []
                    source1 = str(row[6])
                    if source1 != 'nan':
                        if source1 != '':
                            sources += source1.split('\n')
                            source1 = str(source1.replace("\n", ""))
                            source1 = str(source1.replace(" ", ""))
                            source1 = str(source1.replace("NaN", ""))
                            source1 = str(source1.replace("nan", ""))
                    sources = list(filter(None, sources))
                    #print(sources)

                    #get cell destinations
                    #destinations = get_destinations(data)
                    destinations = []
                    destination1 = str(row[9])
                    if destination1 != 'nan':
                        if destination1 != '':
                            destinations += destination1.split('\n')
                            destination1 = str(destination1.replace("\n", ""))
                            destination1 = str(destination1.replace(" ", ""))
                            destination1 = str(destination1.replace("NaN", ""))
                            destination1 = str(destination1.replace("nan", ""))
                    destinations = list(filter(None, destinations))        
                    #print(destinations)

                    #get services from cell
                    #services = get_service_objects(data)
                    services = []
                    service1 = str(row[12])
                    if service1 != 'nan':
                        services += service1.split('\n')
                        service1 = str(service1.replace("\n", ""))
                        service1 = str(service1.replace(" ", ""))
                    services = list(filter(None, services))
                    #print(services)

                    #get description from cell
                    description = str(row[16])
                    description = f"{description}\n{session_name}"
                    #todo!
                    #name = get_rule_name(section_name, section_id=auth_header)
                    #print(description)
                    section = get_section_name(section_name)
                    name = new_description = input(f"Please enter the name for the new rule in section {section}: ")
                    

                    url = f"https://{mgmt_ip}/web_api/add-access-rule"
                    data = {
                        "layer": "Network",
                        "position": {
                            "bottom": section_id
                        },
                        "name": name,
                        "source": sources,
                        "destination": destinations,
                        "service": services,
                        "action": "accept",
                        "comments": description,
                        "track": "log",
                        "install-on": install_on
                    }

                    response = requests.post(url, headers=auth_header, data=json.dumps(data), verify=False)
                    # Check the response
                    if response.status_code == 200:
                        print("[+] Rule created successfully")
                    else:
                        print(f"[!] Error creating rule: {response.status_code} {response.text}")
                        discard_changes(auth_header)
                else:
                    print("[!] Unknown section")


            elif rule_type == "CHANGE" or rule_type == "change" or rule_type == "chg" or rule_type == "CHG" or rule_type == "Change" or rule_type == "Chg":
                #print("Rule will be changed")
                rule_id = str(row[4])
                rule_id = str(rule_id.replace("\n", ""))
                rule_id = str(rule_id.replace(" ", ""))
                if rule_id:
                    print(f"[*] Rule will be changed: {rule_id}")
                    rule_uid = get_rule_id(auth_header, rule_id)
                    if rule_uid:
                        #print(rule_uid)

                        #check source changes
                        sources = []
                        source1 = str(row[6])
                        if source1 != 'nan':
                            if source1 != '':
                                sources += source1.split('\n')
                                source1 = str(source1.replace("\n", ""))
                                source1 = str(source1.replace(" ", ""))
                                source1 = str(source1.replace("NaN", ""))
                                source1 = str(source1.replace("nan", ""))
                        sources = list(filter(None, sources))
                        #print(sources)
                        add_sources = []
                        del_sources = []
                        if sources:
                            action = ""
                            for source in sources:
                                if source == "Add" or source == "add" or source == "ADD" or source == "Add:" or source == "add:" or source == "ADD:":
                                    action = "add"
                                elif source == "Del" or source == "del" or source == "DEL" or source == "Del:" or source == "del:" or source == "DEL:" or source == "delete" or source == "delete:" or source == "Delete" or source == "Delete:" or source == "DELETE" or source == "DELETE:":
                                    action = "delete"
                                else:
                                    if action == "add":
                                        add_sources.append(source)
                                    elif action == "delete":
                                        del_sources.append(source)
                                    else:
                                        print(f"No action set, nothing happens with {source}")                 
                            print(f"[*] {add_sources} will be added to rule (source)...")
                            print(f"[*] {del_sources} will be removed from rule (source)...")
                            layer = "Network"
                            if add_sources:
                                payload = {
                                    "uid": rule_uid,
                                    "layer": layer,
                                    "source":{
                                        "add": add_sources
                                    }
                                }
                                url = f"https://{mgmt_ip}/web_api/set-access-rule"
                                response = requests.post(url, headers=auth_header, data=json.dumps(payload), verify=False)
                                # Check the response
                                if response.status_code == 200:
                                    print("[+] Successfully added sources to rule")
                                else:
                                    print(f"[!] Error adding source to rule: {response.status_code} {response.text}")
                                    discard_changes(auth_header)
                            if del_sources:
                                payload = {
                                    "uid": rule_uid,
                                    "layer": layer,
                                    "source":{
                                        "remove": del_sources
                                    }
                                }
                                url = f"https://{mgmt_ip}/web_api/set-access-rule"
                                response = requests.post(url, headers=auth_header, data=json.dumps(payload), verify=False)
                                # Check the response
                                if response.status_code == 200:
                                    print("[+] Successfully removed sources from rule")
                                else:
                                    print(f"[!] Error removing sources from rule: {response.status_code} {response.text}")
                                    discard_changes(auth_header)

                        #check destination changes
                        destinations = []
                        destination1 = str(row[9])
                        if destination1 != 'nan':
                            if destination1 != '':
                                destinations += destination1.split('\n')
                                destination1 = str(destination1.replace("\n", ""))
                                destination1 = str(destination1.replace(" ", ""))
                                destination1 = str(destination1.replace("NaN", ""))
                                destination1 = str(destination1.replace("nan", ""))
                        destinations = list(filter(None, destinations))
                        #print(destinations)
                        add_destinations = []
                        del_destinations = []
                        if destinations:
                            action = ""
                            for destination in destinations:
                                if destination == "Add" or destination == "add" or destination == "ADD" or destination == "Add:" or destination == "add:" or destination == "ADD:":
                                    action = "add"
                                elif destination == "Del" or destinations == "del" or destination == "DEL" or destination == "Del:" or destination == "del:" or destination == "DEL:" or destination == "delete" or destination == "delete:" or destination == "Delete" or destination == "Delete:" or destination == "DELETE" or destination == "DELETE:":
                                    action = "delete"
                                else:
                                    if action == "add":
                                        add_destinations.append(destination)
                                    elif action == "delete":
                                        del_destinations.append(destination)
                                    else:
                                        print(f"[*] No action set, nothing happens with {destination}")                 
                            print(f"[*] {add_destinations} will be added to rule...")
                            print(f"[*] {del_destinations} will be removed from rule...")
                            layer = "Network"
                            if add_destinations:
                                payload = {
                                    "uid": rule_uid,
                                    "layer": layer,
                                    "destination":{
                                        "add": add_destinations
                                    }
                                }
                                url = f"https://{mgmt_ip}/web_api/set-access-rule"
                                response = requests.post(url, headers=auth_header, data=json.dumps(payload), verify=False)
                                # Check the response
                                if response.status_code == 200:
                                    print("[+] Successfully added destinations to rule")
                                else:
                                    print(f"[!] Error adding destination to rule: {response.status_code} {response.text}")
                                    discard_changes(auth_header)
                            if del_destinations:
                                payload = {
                                    "uid": rule_uid,
                                    "layer": layer,
                                    "source":{
                                        "remove": del_destinations
                                    }
                                }
                                url = f"https://{mgmt_ip}/web_api/set-access-rule"
                                response = requests.post(url, headers=auth_header, data=json.dumps(payload), verify=False)
                                # Check the response
                                if response.status_code == 200:
                                    print("[+] Successfully removed destinations from rule")
                                else:
                                    print(f"[!] Error removing destinations from rule: {response.status_code} {response.text}")
                                    discard_changes(auth_header)

                        # check install on changes
                        install_on = []
                        install_on1 = str(row[1])
                        if install_on1 != 'nan':
                            if install_on1 != '':
                                install_on += install_on1.split('\n')
                                install_on1 = str(install_on1.replace("\n", ""))
                                install_on1 = str(install_on1.replace(" ", ""))
                                install_on1 = str(install_on1.replace("NaN", ""))
                                install_on1 = str(install_on1.replace("nan", ""))
                        install_on = list(filter(None, install_on))
                        #print(install_on)
                        add_install_on = []
                        del_install_on = []
                        if install_on:
                            action = ""
                            for obj in install_on:
                                if obj == "Add" or obj == "add" or obj == "ADD" or obj == "Add:" or obj == "add:" or obj == "ADD:":
                                    action = "add"
                                elif obj == "Del" or obj == "del" or obj == "DEL" or obj == "Del:" or obj == "del:" or obj == "DEL:" or obj == "delete" or obj == "delete:" or obj == "Delete" or obj == "Delete:" or obj == "DELETE" or obj == "DELETE:":
                                    action = "delete"
                                else:
                                    if action == "add":
                                        add_install_on.append(obj)
                                    elif action == "delete":
                                        del_install_on.append(obj)
                                    else:
                                        print(f"[*] No action set, nothing happens with {obj}")                 
                            print(f"[*] {add_install_on} will be added to rule...")
                            print(f"[*] {del_install_on} will be removed from rule...")
                            layer = "Network"
                            if add_install_on:
                                payload = {
                                    "uid": rule_uid,
                                    "layer": layer,
                                    "install_on":{
                                        "add": add_install_on
                                    }
                                }
                                url = f"https://{mgmt_ip}/web_api/set-access-rule"
                                response = requests.post(url, headers=auth_header, data=json.dumps(payload), verify=False)
                                # Check the response
                                if response.status_code == 200:
                                    print("[+] Successfully added installation gateways to rule")
                                else:
                                    print(f"[!] Error adding installation gateways to rule: {response.status_code} {response.text}")
                                    discard_changes(auth_header)
                            if del_install_on:
                                payload = {
                                    "uid": rule_uid,
                                    "layer": layer,
                                    "install_on":{
                                        "remove": del_install_on
                                    }
                                }
                                url = f"https://{mgmt_ip}/web_api/set-access-rule"
                                response = requests.post(url, headers=auth_header, data=json.dumps(payload), verify=False)
                                # Check the response
                                if response.status_code == 200:
                                    print("[+] Successfully removed installation gateway from rule")
                                else:
                                    print(f"[!] Error removing installation gateway from rule: {response.status_code} {response.text}")
                                    discard_changes(auth_header)

                        # check service changes
                        services = []
                        service1 = str(row[12])
                        if service1 != 'nan':
                            if service1 != '':
                                services += service1.split('\n')
                                service1 = str(service1.replace("\n", ""))
                                service1 = str(service1.replace(" ", ""))
                                service1 = str(service1.replace("NaN", ""))
                                service1 = str(service1.replace("nan", ""))
                        services = list(filter(None, services))
                        #print(services)
                        add_services = []
                        del_services = []
                        if services:
                            action = ""
                            for obj in services:
                                if obj == "Add" or obj == "add" or obj == "ADD" or obj == "Add:" or obj == "add:" or obj == "ADD:":
                                    action = "add"
                                elif obj == "Del" or obj == "del" or obj == "DEL" or obj == "Del:" or obj == "del:" or obj == "DEL:" or obj == "delete" or obj == "delete:" or obj == "Delete" or obj == "Delete:" or obj == "DELETE" or obj == "DELETE:":
                                    action = "delete"
                                else:
                                    if action == "add":
                                        add_services.append(obj)
                                    elif action == "delete":
                                        del_services.append(obj)
                                    else:
                                        print(f"[*] No action set, nothing happens with {obj}")                 
                            print(f"[*] {add_services} will be added to rule...")
                            print(f"[*] {del_services} will be removed from rule...")
                            layer = "Network"
                            if add_services:
                                payload = {
                                    "uid": rule_uid,
                                    "layer": layer,
                                    "service":{
                                        "add": add_services
                                    }
                                }
                                url = f"https://{mgmt_ip}/web_api/set-access-rule"
                                response = requests.post(url, headers=auth_header, data=json.dumps(payload), verify=False)
                                # Check the response
                                if response.status_code == 200:
                                    print("[+] Successfully added services to rule")
                                else:
                                    print(f"[!] Error adding services to rule: {response.status_code} {response.text}")
                                    discard_changes(auth_header)
                            if del_services:
                                payload = {
                                    "uid": rule_uid,
                                    "layer": layer,
                                    "service":{
                                        "remove": del_services
                                    }
                                }
                                url = f"https://{mgmt_ip}/web_api/set-access-rule"
                                response = requests.post(url, headers=auth_header, data=json.dumps(payload), verify=False)
                                # Check the response
                                if response.status_code == 200:
                                    print("[+] Successfully removed services from rule")
                                else:
                                    print(f"[!] Error removing services from rule: {response.status_code} {response.text}")
                                    discard_changes(auth_header)


                        # change description
                        description = get_description(auth_header, rule_uid)
                        new_description = f"{description}\n{session_name}"
                        if new_description:
                            payload = {
                                "uid": rule_uid,
                                "layer": layer,
                                "comments": new_description
                            }
                            url = f"https://{mgmt_ip}/web_api/set-access-rule"
                            response = requests.post(url, headers=auth_header, data=json.dumps(payload), verify=False)
                            # Check the response
                            if response.status_code == 200:
                                print("[+] Successfully changed description")
                            else:
                                print(f"[!] Error adding description: {response.status_code} {response.text}")
                                discard_changes(auth_header)


                        
                    else:
                        print("[!] Unknown rule uid")
            elif rule_type == "DELETE" or rule_type == "delete" or rule_type =="DEL" or rule_type == "del" or rule_type == "Delete" or rule_type == "Del":
                rule_id = str(row[4])
                rule_id = str(rule_id.replace("\n", ""))
                rule_id = str(rule_id.replace(" ", ""))
                if rule_id:
                    #print(rule_id)
                    rule_uid = get_rule_id(auth_header, rule_id)
                    print(f"[*] Deleting rule: {rule_uid}")
                    if rule_uid:
                        layer = "Network"
                        payload = {
                            "uid": rule_uid,
                            "layer": layer
                        }
                        url = f"https://{mgmt_ip}/web_api/delete-access-rule"
                        response = requests.post(url, headers=auth_header, data=json.dumps(payload), verify=False)
                        # Check the response
                        if response.status_code == 200:
                            print("[+] Rule deleted successfully")
                        else:
                            print(f"[!] Error deleting rule: {response.status_code} {response.text}")
                            discard_changes(auth_header)

                else:
                    print("[!] No Rule_ID has been found!")



def change_session_information(auth_header):
    today = datetime.date.today()
    today_str = today.strftime("%d.%m.%Y")
    session_id = auth_header['X-chkp-sid']
    new_name = f"api_user@{today_str}"
    new_description = input("[*] Please enter a description for the session: ")
    payload = {
        "new-name": new_name,
        "description": new_description
    }
    data = json.dumps(payload)
    url = f"https://{mgmt_ip}/web_api/set-session"
    response = requests.post(url, data=data, headers=auth_header, verify=False)
    if response.status_code == 200:
        print("[+] Session has been renamed")
        return new_description
    else:
        print(f"[!] Error renaming session: {response.text}")
        discard_changes(auth_header)


def discard_changes(auth_header):
    print("[!] Error - discarding changes...")
    url = f"https://{mgmt_ip}/web_api/discard"
    discard = {}
    data = json.dumps(discard)
    response = requests.post(url, data=data, headers=auth_header, verify=False)
    if response.status_code != 200:
        print(f"[!] Discarding changes failed with error: {response.text}")
        sys.exit()
    else:
        print("[+] Changes discarded successfully.")
        sys.exit()

def publish_changes(auth_header):
    url = f"https://{mgmt_ip}/web_api/publish"
    publish = {}
    data = json.dumps(publish)
    response = requests.post(url, data=data, headers=auth_header, verify=False)
    if response.status_code != 200:
        print(f"[!] Publishing changes failed with error: {response.text}")
        discard_changes(auth_header)
    else:
        print("[+] Changes published successfully.")

def get_session_id(auth_header):
    url = f"https://{mgmt_ip}/web_api/show-session"
    payload = {}
    data = json.dumps(payload)
    response = requests.post(url, data=data, headers=auth_header, verify=False)
    response_json = json.loads(response.text)
    session_uid = response_json['uid']
    return session_uid

def submit_session(auth_header):
    url = f"https://{mgmt_ip}/web_api/submit-session"
    session_uid = get_session_id(auth_header)
    payload = {
        "uid": session_uid
    }
    data = json.dumps(payload)
    response = requests.post(url, data=data, headers=auth_header, verify=False)
    if response.status_code != 200:
        print(f"[!] Submitting changes failed with error: {response.text}")
        discard_changes(auth_header)
    else:
        print("[+] Changes submitted successfully.")


def disconnect_session(auth_header):
    url = f"https://{mgmt_ip}/web_api/disconnect"
    session_id = get_session_id(auth_header)
    payload = {
        "uid": session_id
    }
    data = json.dumps(payload)
    response = requests.post(url, data=data, headers=auth_header, verify=False)
    if response.status_code != 200:
        print(f"[!] Disconnecting the session failed: {response.text}")
    else:
        print("[+] Session successfully disconnected.")




def main():
    print_banner()
    # Verbindung zur Check Point API herstellen
    auth_header = connect_checkpoint_api()
    
    # Wähle die Excel-Datei aus
    excel_file = select_excel_file()

    # Extrahiere die Daten aus der Excel-Datei
    data = extract_data(excel_file)

    print(f"[*] Data: {data}")

    if auth_header is not None:
        # Name der Session ändern
        session_name = change_session_information(auth_header)

        # Firewall-Objekte extrahieren
        firewall_objects = get_firewall_objects(data)

        # Überprüfen, ob Objekte bereits vorhanden sind
        new_objects = check_objects_exist(auth_header, firewall_objects)

        
        # Anlegen der Objekte
        for obj in new_objects:
            obj = str(obj.replace("\n", ""))
            print(f"[*] Object does not exist: {obj}")
            create_object(auth_header, obj)

        # Service-Objekte extrahieren
        service_objects = get_service_objects(data)

        # Überprüfung, ob Service-Objekte bereits vorhanden sind
        new_services = check_services_exist(auth_header, service_objects)

        #Anlage der Service Objekte
        for obj in new_services:
            obj = str(obj.replace("\n", ""))
            print(f"[*] Object does not exist: {obj}")
            create_service(auth_header, obj)

        #Anlegen der Regeln
        create_change_rules(auth_header, data, session_name)

        #Änderungen publishen --> dies sollte zuletzt in der main()-Funktion durchgeführt werden!
        #Normaler Publish
        #publish_changes(auth_header)

        #Submit Session
        submit_session(auth_header)
        print("[*] Waiting for Publish to finish...")
        time.sleep(10)

        disconnect_session(auth_header)


        #to-do!!!!!
        #Optimierung von Services
        #Optimierung verschiedener Excel Formate
        #check der access role objekte am geclonten management auf r81.20
        #check der VPN access role objekte am geclonten management auf r81.20
        #time objekte hinzufügen
        #andere Policies supporten
        #VPN Rules supporten
        
if __name__ == "__main__":
    main()

