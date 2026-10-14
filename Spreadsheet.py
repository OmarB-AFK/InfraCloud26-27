import json
import csv
import openpyxl
import psutil
import platform
import socket
import subprocess

import datetime
now = datetime.datetime.now()
print(now)


def get_model():
    # Computermodel opvragen (werkt anders op Windows en Linux)
    try:
        if platform.system() == "Windows":
            result = subprocess.check_output(
                'powershell "(Get-CimInstance Win32_ComputerSystem).Model"',
                shell=True, text=True)
            return result.strip()
        else:
            with open("/sys/devices/virtual/dmi/id/product_name") as f:
                return f.read().strip()
    except Exception:
        return None


def get_primary_ip():
    # Zoekt het IP-adres dat gebruikt wordt om naar buiten te verbinden
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))  # er wordt niets verstuurd
        return s.getsockname()[0]
    except Exception:
        return None
    finally:
        s.close()


def collect_system_info():
    data = {}

    # Computer info
    data["computer"] = {
        "computernaam": socket.gethostname(),
        "computermodel": get_model(),
    }

    # CPU info
    freq = psutil.cpu_freq()
    data["cpu"] = {
        "processor": platform.processor(),
        "physical_cores": psutil.cpu_count(logical=False),
        "max_freq_mhz": freq.max if freq else None,
    }

    # Memory info
    data["memory"] = {
        "total_ram_bytes": psutil.virtual_memory().total,
    }

    # Storage info (alle partities)
    storage = {}
    for part in psutil.disk_partitions(all=False):
        try:
            usage = psutil.disk_usage(part.mountpoint)
            storage[part.mountpoint] = {
                "fstype": part.fstype,
                "total_bytes": usage.total,
            }
        except PermissionError:
            continue
    data["storage"] = storage

    # Network info (primaire interface)
    primary_ip = get_primary_ip()
    network = {}
    addrs = psutil.net_if_addrs()
    for iface, addr_list in addrs.items():
        # Is dit de interface met het primaire IPv4-adres?
        is_primary = False
        for a in addr_list:
            if a.family == socket.AF_INET and a.address == primary_ip:
                is_primary = True

        if is_primary:
            network["interface"] = iface
            network["ipv6"] = None
            network["ipv6_link_local"] = None
            for a in addr_list:
                if a.family == psutil.AF_LINK:
                    network["mac"] = a.address
                elif a.family == socket.AF_INET:
                    network["ipv4"] = a.address
                elif a.family == socket.AF_INET6:
                    adres = a.address.split("%")[0]  # %interface weghalen
                    if adres.startswith("fe80"):
                        network["ipv6_link_local"] = adres
                    else:
                        network["ipv6"] = adres
    data["network"] = network

    return data


if __name__ == "__main__":
    info = collect_system_info()

    # Dictionary omzetten naar JSON (tekst)
    json_data = json.dumps(info, indent=2)
    print(json_data)

    # JSON (tekst) terug omzetten naar een dictionary
    dict_data = json.loads(json_data)
    print(dict_data)

    # Dictionary omzetten naar een CSV-bestand
    kolommen = [
        "computernaam", "computermodel", "processor", "fysieke_cores",
        "max_freq_mhz", "ram_bytes", "mountpoint", "bestandssysteem",
        "opslag_bytes", "mac", "ipv4", "ipv6", "ipv6_link_local",
    ]

    with open("systeminfo.csv", "w", newline="") as bestand:
        schrijver = csv.writer(bestand)
        schrijver.writerow(kolommen)  # eerste rij = titels

        # Eén rij per partitie
        for mountpoint in dict_data["storage"]:
            schijf = dict_data["storage"][mountpoint]
            rij = [
                dict_data["computer"]["computernaam"],
                dict_data["computer"]["computermodel"],
                dict_data["cpu"]["processor"],
                dict_data["cpu"]["physical_cores"],
                dict_data["cpu"]["max_freq_mhz"],
                dict_data["memory"]["total_ram_bytes"],
                mountpoint,
                schijf["fstype"],
                schijf["total_bytes"],
                dict_data["network"].get("mac"),
                dict_data["network"].get("ipv4"),
                dict_data["network"].get("ipv6"),
                dict_data["network"].get("ipv6_link_local"),
            ]
            schrijver.writerow(rij)

    # Dictionary omzetten naar een Excel-bestand (.xlsx)
    werkboek = openpyxl.Workbook()
    blad = werkboek.active
    blad.title = "Systeminfo"

    blad.append(kolommen)  # eerste rij = dezelfde titels als in de CSV

    # Eén rij per partitie
    for mountpoint in dict_data["storage"]:
        schijf = dict_data["storage"][mountpoint]
        rij = [
            dict_data["computer"]["computernaam"],
            dict_data["computer"]["computermodel"],
            dict_data["cpu"]["processor"],
            dict_data["cpu"]["physical_cores"],
            dict_data["cpu"]["max_freq_mhz"],
            dict_data["memory"]["total_ram_bytes"],
            mountpoint,
            schijf["fstype"],
            schijf["total_bytes"],
            dict_data["network"].get("mac"),
            dict_data["network"].get("ipv4"),
            dict_data["network"].get("ipv6"),
            dict_data["network"].get("ipv6_link_local"),
        ]
        blad.append(rij)

    werkboek.save("systeminfo.xlsx")