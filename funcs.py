import wmi
import platform
import psutil
import winreg
import subprocess

from datetime import datetime


def format_date_readable(date):
    parsed = datetime.strptime(date, "%d.%m.%Y")

    months = [
        "января", "февраля", "марта", "апреля", "мая", "июня",
        "июля", "августа", "сентября", "октября", "ноября", "декабря"
    ]

    day = f'"{parsed.day:02d}"'
    month = months[parsed.month - 1]
    year = parsed.year

    return f'{day} {month} {year} г.'

class SystemInfo():
    def get_os_info():
        os_name = platform.system()
        os_version = platform.release()

        edition = "Unknown Edition"

        try:
            key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Windows NT\CurrentVersion"
            )
            product_name, _ = winreg.QueryValueEx(key, "ProductName")
            edition = product_name
        except Exception as e:
            print(f"Не удалось получить редакцию Windows: {e}")

        return f"{edition} ({os_version})"

    def get_cpu_info():
        c = wmi.WMI()
        cpu_info = c.Win32_Processor()[0]
        name = cpu_info.Name.strip()
        freq = psutil.cpu_freq()
        physical_cores = psutil.cpu_count(logical=False)
        logical_cores = psutil.cpu_count(logical=True)
        base_freq = f"{freq.max:.2f} MHz" if freq and freq.max else "Unknown"
        return f"{name} ({physical_cores} ядер / {logical_cores} потоков, до {base_freq})"

    def get_ram_modules_info():
        c = wmi.WMI()

        MEMORY_TYPE_MAP = {
            0: "Unknown", 1: "Other", 2: "DRAM", 3: "Synchronous DRAM", 4: "Cache DRAM",
            5: "EDO", 6: "EDRAM", 7: "VRAM", 8: "SRAM", 9: "RAM", 10: "ROM", 11: "Flash",
            12: "EEPROM", 13: "FEPROM", 14: "EPROM", 15: "CDRAM", 16: "3DRAM", 17: "SDRAM",
            18: "SGRAM", 19: "RDRAM", 20: "DDR", 21: "DDR2", 22: "DDR2 FB-DIMM", 24: "DDR3",
            25: "FBD2", 26: "DDR4", 27: "LPDDR", 28: "LPDDR2", 29: "LPDDR3", 30: "LPDDR4", 31: "LPDDR5",
            32: "HBM", 33: "HBM2", 34: "DDR5", 35: "LPDDR5"
        }

        FORM_FACTOR_MAP = {
            0: "Unknown", 1: "Other", 2: "SIP", 3: "DIP", 4: "ZIP", 5: "SOJ", 6: "Proprietary",
            7: "SIMM", 8: "DIMM", 9: "TSOP", 10: "PGA", 11: "RIMM", 12: "SODIMM", 13: "SRIMM",
            14: "SMD", 15: "SSMP", 16: "QFP", 17: "TQFP", 18: "SOIC", 19: "LCC", 20: "PLCC",
            21: "BGA", 22: "FPBGA", 23: "LGA"
        }

        modules = []
        for mem in c.Win32_PhysicalMemory():
            size_gb = int(getattr(mem, 'Capacity', 0)) // (1024 ** 3)

            smbios_type = getattr(mem, 'SMBIOSMemoryType', 0)
            mem_type = MEMORY_TYPE_MAP.get(smbios_type, "Unknown")

            if mem_type == "Unknown":
                fallback_type = getattr(mem, 'MemoryType', 0)
                mem_type = MEMORY_TYPE_MAP.get(fallback_type, "Unknown")

            form_factor_code = getattr(mem, 'FormFactor', 0)
            form_factor = FORM_FACTOR_MAP.get(form_factor_code, "Unknown")

            modules.append(f"{size_gb} ГБ ({form_factor}, {mem_type})")

        return modules

    def get_serial_number():
        try:
            c = wmi.WMI()
            bios_info = c.Win32_BIOS()[0]
            serial_number = bios_info.SerialNumber.strip()

            if serial_number and serial_number.lower() != "to be filled by o.e.m.":
                return serial_number

            import shlex
            cmd = shlex.split('wmic csproduct get identifyingnumber')
            result = subprocess.check_output(cmd, text=True).splitlines()

            if len(result) >= 2:
                serial_candidate = result[1].strip()
                if serial_candidate and serial_candidate.lower() != "to be filled by o.e.m.":
                    return serial_candidate

            return "Неизвестен"
        except Exception as e:
            return f"Ошибка: {str(e)}"

    def get_all_system_info():
        c = wmi.WMI()

        # model
        system_info = c.Win32_ComputerSystem()[0]
        bios_info = c.Win32_BIOS()[0]
        model = system_info.Model

        # s\n
        serial_number = SystemInfo.get_serial_number()

        # OS
        os_info = SystemInfo.get_os_info()

        # CPU
        cpu = SystemInfo.get_cpu_info()

        # RAM
        ram_total = (f'{round(psutil.virtual_memory().total / (1024 ** 3), 2)}')

        # RAM modules
        memory_info = SystemInfo.get_ram_modules_info()
        ram_modules_formatted = "\n".join(f"- {module}" for module in memory_info)

        # drives
        drives = []
        for disk in c.Win32_DiskDrive():
            size_gb = int(disk.Size) // (1024 ** 3)
            drives.append((disk.Model.strip(), f"{size_gb} GB"))
        drives_formatted = "\n".join(f"- {name} — {size}" for name, size in drives)

        return {
            "model": model,
            "serial": serial_number,
            "OS": os_info,
            "CPU": cpu,
            "RAM": f"{ram_total} GB",
            "RAM_TYPE": ram_modules_formatted,
            "drives": drives_formatted,
        }

    def replace_placeholders(doc, replacements):
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                for key, val in replacements.items():
                    if key in run.text:
                        run.text = run.text.replace(key, str(val))

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            for key, val in replacements.items():
                                if key in run.text:
                                    run.text = run.text.replace(key, str(val))