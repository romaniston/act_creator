import wmi
import platform
import psutil
from docx import Document

import os


def get_system_info():
    c = wmi.WMI()

    # model & serial
    system_info = c.Win32_ComputerSystem()[0]
    bios_info = c.Win32_BIOS()[0]
    model = system_info.Model
    serial_number = bios_info.SerialNumber.strip()

    # OS
    os_info = platform.platform()

    # CPU
    cpu = platform.processor()

    # RAM
    ram_total = round(psutil.virtual_memory().total / (1024 ** 3), 2)

    # RAM modules
    memory_info = []
    for mem in c.Win32_PhysicalMemory():
        mem_size = int(mem.Capacity) // (1024 ** 3)
        mem_type = mem.MemoryType
        memory_info.append((mem.Manufacturer.strip(), f"{mem_size} GB", f"Type: {mem.MemoryType}"))

    # drives
    drives = []
    for disk in c.Win32_DiskDrive():
        size_gb = int(disk.Size) // (1024 ** 3)
        drives.append((disk.Model.strip(), f"{size_gb} GB"))

    return {
        "model": model,
        "serial": serial_number,
        "OS": os_info,
        "CPU": cpu,
        "RAM": f"{ram_total} GB",
        "RAM modules": memory_info,
        "drives": drives,
    }


if __name__ == "__main__":
    from pprint import pprint
    pprint(get_system_info())
