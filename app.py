import os
import subprocess
import sys
from datetime import datetime

import psutil
import wmi
from PyQt5.QtCore import QRectF, Qt
from PyQt5.QtGui import QColor, QFont, QFontMetrics, QIcon, QPainter, QPen, QPixmap
from PyQt5.QtSvg import QSvgRenderer
from PyQt5.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QGridLayout,
    QGroupBox,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)


# ------------------------------
# Utility formatting helpers
# ------------------------------
def format_bytes(value):
    if value is None:
        return "N/A"
    units = ["B", "KB", "MB", "GB", "TB", "PB"]
    size = float(value)
    for unit in units:
        if size < 1024.0 or unit == units[-1]:
            return f"{size:.2f} {unit}"
        size /= 1024.0
    return f"{size:.2f} PB"


def format_marketed_storage(value_bytes):
    if value_bytes is None:
        return "N/A"
    try:
        size = float(value_bytes)
    except Exception:
        return "N/A"
    if size <= 0:
        return "N/A"

    decimal_tb = size / (10**12)
    decimal_gb = size / (10**9)

    if decimal_tb >= 1:
        rounded_tb = int(round(decimal_tb))
        if abs(decimal_tb - rounded_tb) <= 0.15:
            return f"{rounded_tb} TB"
        return f"{decimal_tb:.1f} TB"

    rounded_gb = int(round(decimal_gb))
    return f"{rounded_gb} GB"


def safe_text(value, fallback="N/A"):
    if value is None:
        return fallback
    value = str(value).strip()
    return value if value else fallback


def get_logo_path():
    logo_path = os.path.join(os.path.dirname(__file__), "img", "Logo.png")
    return logo_path if os.path.exists(logo_path) else None


def get_section_icon_path(title, os_name=None):
    icon_map = {
        "CPU": "processor.svg",
        "GPU": "graphics-card.png",
        "RAM": "RAM.svg",
        "Storage": "storage.svg",
        "Motherboard": "motherboard.png",
        "BIOS": "bios.png",
    }

    if title == "Operating System":
        normalized_os = str(os_name or "").lower()
        if "linux" in normalized_os:
            file_name = "linux-platform.png"
        elif "mac" in normalized_os or "darwin" in normalized_os or "os x" in normalized_os:
            file_name = "apple-logo.png"
        else:
            file_name = "windows.png"
    else:
        file_name = icon_map.get(title)

    if not file_name:
        return None

    icon_path = os.path.join(os.path.dirname(__file__), "img", file_name)
    return icon_path if os.path.exists(icon_path) else None


def _is_likely_integrated_gpu_name(gpu_name):
    text = str(gpu_name or "").lower()
    integrated_markers = [
        "intel",
        "uhd",
        "iris",
        "hd graphics",
        "integrated",
        "apu",
        "radeon graphics",
    ]
    return any(marker in text for marker in integrated_markers)


def _is_likely_discrete_gpu_name(gpu_name):
    text = str(gpu_name or "").lower()
    discrete_markers = [
        "nvidia",
        "geforce",
        "rtx",
        "gtx",
        "quadro",
        "tesla",
        "titan",
        "radeon rx",
        "radeon pro",
        "intel arc",
        " arc ",
    ]
    return any(marker in text for marker in discrete_markers)


def _gpu_priority_key(gpu):
    name = safe_text(gpu.get("Name", ""), "")
    memory_bytes = gpu.get("_memory_bytes", 0)
    try:
        memory_bytes = int(memory_bytes)
    except Exception:
        memory_bytes = 0
    if memory_bytes < 0:
        memory_bytes = 0

    if _is_likely_discrete_gpu_name(name):
        tier = 3
    elif not _is_likely_integrated_gpu_name(name):
        tier = 2
    else:
        tier = 1

    return (tier, memory_bytes)


def _pick_preferred_gpu(gpus):
    if not gpus:
        return {"Name": "N/A", "Memory": "N/A"}
    return max(gpus, key=_gpu_priority_key)


def _gpu_type_label(gpu_name):
    if _is_likely_discrete_gpu_name(gpu_name):
        return "External (Discrete)"
    if _is_likely_integrated_gpu_name(gpu_name):
        return "Integrated"
    return "Unknown"


def _query_nvidia_smi_gpus():
    try:
        result = subprocess.run(
            [
                "nvidia-smi",
                "--query-gpu=name,memory.total,driver_version",
                "--format=csv,noheader,nounits",
            ],
            capture_output=True,
            text=True,
            check=True,
            timeout=3,
        )
    except Exception:
        return []

    records = []
    for raw_line in (result.stdout or "").splitlines():
        line = raw_line.strip()
        if not line:
            continue
        parts = [part.strip() for part in line.split(",")]
        if len(parts) < 3:
            continue
        name, memory_mb_text, driver_version = parts[0], parts[1], parts[2]
        try:
            memory_mb = int(float(memory_mb_text))
        except Exception:
            memory_mb = 0
        records.append(
            {
                "name": name,
                "memory_mb": memory_mb,
                "driver_version": driver_version,
            }
        )
    return records


# ------------------------------
# Hardware/system information collectors
# ------------------------------
def get_cpu_info(wmi_client):
    cpu_data = {
        "Name": "N/A",
        "Physical Cores": "N/A",
        "Logical Threads": "N/A",
        "Current Frequency": "N/A",
        "Max Frequency": "N/A",
    }

    try:
        cpus = wmi_client.Win32_Processor()
        if cpus:
            cpu = cpus[0]
            cpu_data["Name"] = safe_text(cpu.Name)
    except Exception:
        pass

    try:
        cpu_data["Physical Cores"] = str(psutil.cpu_count(logical=False) or "N/A")
        cpu_data["Logical Threads"] = str(psutil.cpu_count(logical=True) or "N/A")
    except Exception:
        pass

    try:
        freq = psutil.cpu_freq()
        if freq:
            cpu_data["Current Frequency"] = f"{freq.current:.2f} MHz"
            cpu_data["Max Frequency"] = f"{freq.max:.2f} MHz" if freq.max else "N/A"
    except Exception:
        pass

    return cpu_data


def get_gpu_info(wmi_client):
    gpu_candidates = []
    nvidia_smi_gpus = _query_nvidia_smi_gpus()

    try:
        gpus = wmi_client.Win32_VideoController()
        for gpu in gpus:
            memory_value = None
            try:
                if gpu.AdapterRAM:
                    memory_value = int(gpu.AdapterRAM)
            except Exception:
                memory_value = None

            normalized_memory = memory_value if (memory_value and memory_value > 0) else 0

            name = safe_text(getattr(gpu, "Name", None))
            driver_version = safe_text(getattr(gpu, "DriverVersion", None))

            if normalized_memory == 0 and _is_likely_discrete_gpu_name(name):
                for nvidia_gpu in nvidia_smi_gpus:
                    nvidia_name = nvidia_gpu.get("name", "")
                    if not nvidia_name:
                        continue
                    left = name.lower()
                    right = nvidia_name.lower()
                    if left in right or right in left:
                        memory_mb = nvidia_gpu.get("memory_mb", 0)
                        if memory_mb > 0:
                            normalized_memory = memory_mb * 1024 * 1024
                        smi_driver = safe_text(nvidia_gpu.get("driver_version", None))
                        if smi_driver != "N/A":
                            driver_version = smi_driver
                        break

            gpu_candidates.append(
                {
                    "Name": name,
                    "Memory": format_bytes(normalized_memory) if normalized_memory else "N/A",
                    "_memory_bytes": normalized_memory,
                    "Driver": driver_version,
                    "Type": _gpu_type_label(name),
                }
            )
    except Exception:
        pass

    if not gpu_candidates:
        return [{"Name": "N/A", "Memory": "N/A"}]

    primary_gpu = _pick_preferred_gpu(gpu_candidates)

    ordered = [primary_gpu] + [gpu for gpu in gpu_candidates if gpu is not primary_gpu]

    gpu_list = []
    for gpu in ordered:
        gpu_list.append(
            {
                "Name": gpu["Name"],
                "Memory": gpu["Memory"],
                "Driver": gpu.get("Driver", "N/A"),
                "Type": gpu.get("Type", "Unknown"),
            }
        )

    return gpu_list


def get_ram_info(wmi_client):
    ram_data = {
        "Total": "N/A",
        "Installed": "N/A",
        "Usable": "N/A",
        "Used": "N/A",
        "Modules": "N/A",
        "Module Layout": "N/A",
        "Speed": "N/A",
    }

    try:
        vm = psutil.virtual_memory()
        ram_data["Total"] = format_bytes(vm.total)
        ram_data["Usable"] = format_bytes(vm.total)
        ram_data["Used"] = format_bytes(vm.used)
    except Exception:
        pass

    try:
        modules = wmi_client.Win32_PhysicalMemory()

        capacities = []
        for module in modules:
            try:
                capacity_value = int(getattr(module, "Capacity", 0) or 0)
            except Exception:
                capacity_value = 0
            if capacity_value > 0:
                capacities.append(capacity_value)

        if capacities:
            total_installed_bytes = sum(capacities)
            installed_gb = max(1, int(round(total_installed_bytes / (1024 ** 3))))
            ram_data["Installed"] = f"{installed_gb} GB"
            ram_data["Total"] = f"{installed_gb} GB"
            ram_data["Modules"] = str(len(capacities))

            module_size_gb = []
            for capacity in capacities:
                size_gb = max(1, int(round(capacity / (1024 ** 3))))
                module_size_gb.append(size_gb)

            layout_counts = {}
            for size in module_size_gb:
                layout_counts[size] = layout_counts.get(size, 0) + 1

            layout_parts = []
            for size in sorted(layout_counts):
                count = layout_counts[size]
                layout_parts.append(f"{count} x {size} GB")
            ram_data["Module Layout"] = " + ".join(layout_parts)

        speeds = [int(m.Speed) for m in modules if getattr(m, "Speed", None)]
        if speeds:
            unique_speeds = sorted(set(speeds))
            if len(unique_speeds) == 1:
                ram_data["Speed"] = f"{unique_speeds[0]} MHz"
            else:
                ram_data["Speed"] = ", ".join(f"{s} MHz" for s in unique_speeds)
    except Exception:
        pass

    return ram_data


def get_storage_info(wmi_client):
    logical_total = 0
    logical_free = 0
    file_systems = []
    physical_total = 0

    try:
        logical_disks = wmi_client.Win32_LogicalDisk(DriveType=3)
        for disk in logical_disks:
            try:
                size_value = int(getattr(disk, "Size", 0) or 0)
                free_value = int(getattr(disk, "FreeSpace", 0) or 0)
            except Exception:
                continue

            if size_value <= 0:
                continue

            logical_total += size_value
            logical_free += max(0, min(free_value, size_value))

            fs = safe_text(getattr(disk, "FileSystem", None), "")
            if fs:
                file_systems.append(fs)
    except Exception:
        pass

    try:
        physical_disks = wmi_client.Win32_DiskDrive()
        for disk in physical_disks:
            try:
                disk_size = int(getattr(disk, "Size", 0) or 0)
            except Exception:
                disk_size = 0
            if disk_size > 0:
                physical_total += disk_size
    except Exception:
        pass

    usable_total = logical_total if logical_total > 0 else None
    installed_total = physical_total if physical_total > 0 else usable_total
    used_total = None
    if usable_total is not None:
        used_total = max(0, usable_total - logical_free)

    fs_summary = "N/A"
    if file_systems:
        fs_summary = ", ".join(sorted(set(file_systems)))

    summary = {
        "Drive": "Local Storage",
        "Mount": "All local volumes",
        "File System": fs_summary,
        "Installed": format_marketed_storage(installed_total),
        "Usable": format_bytes(usable_total) if usable_total is not None else "N/A",
        "Used": format_bytes(used_total) if used_total is not None else "N/A",
        "Total": format_bytes(usable_total) if usable_total is not None else "N/A",
        "Free": format_bytes(logical_free) if usable_total is not None else "N/A",
    }

    return [summary]


def get_motherboard_info(wmi_client):
    board_data = {
        "Manufacturer": "N/A",
        "Model": "N/A",
    }

    try:
        boards = wmi_client.Win32_BaseBoard()
        if boards:
            board = boards[0]
            board_data["Manufacturer"] = safe_text(getattr(board, "Manufacturer", None))
            board_data["Model"] = safe_text(getattr(board, "Product", None))
    except Exception:
        pass

    return board_data


def get_bios_info(wmi_client):
    bios_data = {
        "BIOS Version": "N/A",
        "Release Date": "N/A",
    }

    try:
        bios_items = wmi_client.Win32_BIOS()
        if bios_items:
            bios = bios_items[0]
            bios_data["BIOS Version"] = safe_text(getattr(bios, "SMBIOSBIOSVersion", None))
            raw_date = safe_text(getattr(bios, "ReleaseDate", None), "")
            bios_data["Release Date"] = raw_date[:8] if raw_date else "N/A"
    except Exception:
        pass

    return bios_data


def get_os_info(wmi_client):
    os_data = {
        "Name": "N/A",
        "Version": "N/A",
        "Build": "N/A",
        "Architecture": "N/A",
    }

    try:
        systems = wmi_client.Win32_OperatingSystem()
        if systems:
            system = systems[0]
            os_data["Name"] = safe_text(getattr(system, "Caption", None))
            os_data["Version"] = safe_text(getattr(system, "Version", None))
            os_data["Build"] = safe_text(getattr(system, "BuildNumber", None))
            os_data["Architecture"] = safe_text(getattr(system, "OSArchitecture", None))
    except Exception:
        pass

    return os_data


def collect_system_specs():
    wmi_client = wmi.WMI()
    return {
        "CPU": get_cpu_info(wmi_client),
        "GPU": get_gpu_info(wmi_client),
        "RAM": get_ram_info(wmi_client),
        "Storage": get_storage_info(wmi_client),
        "Motherboard": get_motherboard_info(wmi_client),
        "BIOS": get_bios_info(wmi_client),
        "OS": get_os_info(wmi_client),
        "Scanned At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


# ------------------------------
# Export formatter helpers
# ------------------------------
def specs_to_text(specs):
    lines = []
    lines.append("System Hardware Information")
    lines.append("=" * 40)
    lines.append(f"Scanned At: {specs.get('Scanned At', 'N/A')}")
    lines.append("")

    lines.append("[CPU]")
    for key, value in specs["CPU"].items():
        lines.append(f"{key}: {value}")
    lines.append("")

    lines.append("[GPU]")
    for index, gpu in enumerate(specs["GPU"], start=1):
        lines.append(f"GPU {index}:")
        lines.append(f"  Name: {gpu.get('Name', 'N/A')}")
        lines.append(f"  Memory: {gpu.get('Memory', 'N/A')}")
    lines.append("")

    lines.append("[RAM]")
    for key, value in specs["RAM"].items():
        lines.append(f"{key}: {value}")
    lines.append("")

    lines.append("[Storage]")
    for drive in specs["Storage"]:
        lines.append(f"Drive: {drive.get('Drive', 'N/A')}")
        lines.append(f"  Mount: {drive.get('Mount', 'N/A')}")
        lines.append(f"  File System: {drive.get('File System', 'N/A')}")
        lines.append(f"  Total: {drive.get('Total', 'N/A')}")
        lines.append(f"  Free: {drive.get('Free', 'N/A')}")
    lines.append("")

    lines.append("[Motherboard]")
    for key, value in specs["Motherboard"].items():
        lines.append(f"{key}: {value}")
    lines.append("")

    lines.append("[BIOS]")
    for key, value in specs["BIOS"].items():
        lines.append(f"{key}: {value}")
    lines.append("")

    lines.append("[OS]")
    for key, value in specs["OS"].items():
        lines.append(f"{key}: {value}")

    return "\n".join(lines)


def _wrap_text(text, metrics, max_width):
    words = str(text).split()
    if not words:
        return [""]

    lines = []
    current_line = words[0]
    for word in words[1:]:
        candidate = f"{current_line} {word}"
        if metrics.horizontalAdvance(candidate) <= max_width:
            current_line = candidate
        else:
            lines.append(current_line)
            current_line = word
    lines.append(current_line)
    return lines


def _wrap_multiline(text, metrics, max_width):
    wrapped_lines = []
    for paragraph in str(text).splitlines() or [""]:
        wrapped_lines.extend(_wrap_text(paragraph, metrics, max_width))
    return wrapped_lines or [""]


def _get_export_icon_path(title, os_name=None):
    return get_section_icon_path(title, os_name=os_name)


def _build_export_sections(specs):
    cpu = specs.get("CPU", {})
    gpus = specs.get("GPU", [])
    ram = specs.get("RAM", {})
    storage = specs.get("Storage", [])
    board = specs.get("Motherboard", {})
    bios = specs.get("BIOS", {})
    os_info = specs.get("OS", {})

    gpu_lines = []
    for index, gpu in enumerate(gpus, start=1):
        gpu_lines.append(f"GPU {index}: {gpu.get('Name', 'N/A')} ({gpu.get('Memory', 'N/A')})")

    storage_lines = []
    for drive in storage:
        storage_lines.append(
            f"{drive.get('Drive', 'N/A')} - Total: {drive.get('Total', 'N/A')}, Free: {drive.get('Free', 'N/A')}"
        )

    cpu_cores = cpu.get("Physical Cores", "N/A")
    first_gpu = _pick_preferred_gpu(gpus)

    return [
        {
            "title": "CPU",
            "subtitle": f"{cpu.get('Name', 'N/A')}\n{cpu_cores}-core processor",
            "bullets": [
                f"Logical threads: {cpu.get('Logical Threads', 'N/A')}",
                f"Max Frequency: {cpu.get('Max Frequency', 'N/A')}",
            ],
            "icon": _get_export_icon_path("CPU"),
        },
        {
            "title": "GPU",
            "subtitle": first_gpu.get("Name", "N/A"),
            "bullets": [
                f"Memory: {first_gpu.get('Memory', 'N/A')}",
                f"Type: {first_gpu.get('Type', 'Unknown')}",
                f"Driver: {first_gpu.get('Driver', 'N/A')}",
            ],
            "icon": _get_export_icon_path("GPU"),
        },
        {
            "title": "RAM",
            "subtitle": f"Installed: {ram.get('Installed', ram.get('Total', 'N/A'))}",
            "bullets": [
                f"Usable: {ram.get('Usable', 'N/A')}",
                f"Modules: {ram.get('Module Layout', ram.get('Modules', 'N/A'))}",
                f"Speed: {ram.get('Speed', 'N/A')}",
            ],
            "icon": _get_export_icon_path("RAM"),
        },
        {
            "title": "Storage",
            "subtitle": f"Installed: {storage[0].get('Installed', 'N/A') if storage else 'N/A'}",
            "bullets": [
                f"Usable: {storage[0].get('Usable', storage[0].get('Total', 'N/A')) if storage else 'N/A'}",
                f"Used: {storage[0].get('Used', 'N/A') if storage else 'N/A'}",
                f"File System: {storage[0].get('File System', 'N/A') if storage else 'N/A'}",
            ],
            "icon": _get_export_icon_path("Storage"),
        },
        {
            "title": "Motherboard",
            "subtitle": board.get("Model", "N/A"),
            "bullets": [f"Manufacturer: {board.get('Manufacturer', 'N/A')}"],
            "icon": _get_export_icon_path("Motherboard"),
        },
        {
            "title": "BIOS",
            "subtitle": bios.get("BIOS Version", "N/A"),
            "bullets": [f"Release Date: {bios.get('Release Date', 'N/A')}"],
            "icon": _get_export_icon_path("BIOS"),
        },
        {
            "title": "Operating System",
            "subtitle": os_info.get("Name", "N/A"),
            "bullets": [
                f"Version: {os_info.get('Version', 'N/A')}",
                f"Build: {os_info.get('Build', 'N/A')}",
                f"Architecture: {os_info.get('Architecture', 'N/A')}",
            ],
            "icon": _get_export_icon_path("Operating System", os_info.get("Name", "")),
        },
    ]


def _compute_export_card_layout(section, dimensions, metrics):
    card_padding = dimensions["card_padding"]
    card_width = dimensions["card_width"]
    icon_size = dimensions["icon_size"] if section.get("icon") else 0
    icon_gap = dimensions["icon_gap"] if section.get("icon") else 0

    title_metrics = metrics["title"]
    subtitle_metrics = metrics["subtitle"]
    bullet_metrics = metrics["bullet"]

    text_x = card_padding + icon_size + icon_gap
    content_width = card_width - text_x - card_padding

    title_lines = _wrap_multiline(section["title"], title_metrics, content_width)
    subtitle_lines = _wrap_multiline(section["subtitle"], subtitle_metrics, content_width)

    bullet_lines = []
    for bullet in section["bullets"]:
        bullet_lines.extend(_wrap_multiline(f"• {bullet}", bullet_metrics, card_width - (card_padding * 2)))

    title_subtitle_height = (
        len(title_lines) * title_metrics.height()
        + dimensions["title_to_subtitle_gap"]
        + len(subtitle_lines) * subtitle_metrics.height()
    )
    header_height = max(icon_size, title_subtitle_height)
    bullets_height = len(bullet_lines) * bullet_metrics.height()

    card_height = (
        card_padding
        + header_height
        + dimensions["header_to_bullets_gap"]
        + bullets_height
        + card_padding
    )

    return {
        "section": section,
        "title_lines": title_lines,
        "subtitle_lines": subtitle_lines,
        "bullet_lines": bullet_lines,
        "header_height": header_height,
        "height": card_height,
    }


def _draw_icon(painter, icon_path, x, y, size):
    if not icon_path:
        return

    extension = os.path.splitext(icon_path)[1].lower()
    if extension == ".svg":
        renderer = QSvgRenderer(icon_path)
        if renderer.isValid():
            renderer.render(painter, QRectF(float(x), float(y), float(size), float(size)))
        return

    pixmap = QPixmap(icon_path)
    if pixmap.isNull():
        return
    scaled = pixmap.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
    painter.drawPixmap(x, y, scaled)


def export_specs_to_png(specs, output_path):
    sections = _build_export_sections(specs)

    canvas_width = 1080
    canvas_height = 1350

    scale = 1.0
    while True:
        outer_padding = max(16, int(24 * scale))
        column_gap = max(12, int(18 * scale))
        row_gap = max(12, int(18 * scale))
        card_padding = max(14, int(22 * scale))
        icon_size = max(42, int(72 * scale))
        icon_gap = max(8, int(16 * scale))
        card_width = (canvas_width - (outer_padding * 2) - column_gap) // 2

        scan_font = QFont("Ubuntu", max(11, int(14 * scale)))
        title_font = QFont("Ubuntu", max(18, int(30 * scale)), QFont.Bold)
        subtitle_font = QFont("Ubuntu", max(11, int(16 * scale)))
        bullet_font = QFont("Ubuntu", max(10, int(14 * scale)))

        scan_metrics = QFontMetrics(scan_font)
        title_metrics = QFontMetrics(title_font)
        subtitle_metrics = QFontMetrics(subtitle_font)
        bullet_metrics = QFontMetrics(bullet_font)

        dimensions = {
            "card_padding": card_padding,
            "card_width": card_width,
            "icon_size": icon_size,
            "icon_gap": icon_gap,
            "title_to_subtitle_gap": max(4, int(6 * scale)),
            "header_to_bullets_gap": max(8, int(12 * scale)),
        }
        metrics = {
            "title": title_metrics,
            "subtitle": subtitle_metrics,
            "bullet": bullet_metrics,
        }

        rendered = [_compute_export_card_layout(section, dimensions, metrics) for section in sections]
        first_rows = rendered[:6]
        os_row = rendered[6]
        rows = [first_rows[i : i + 2] for i in range(0, len(first_rows), 2)]

        cards_area_height = 0
        for row in rows:
            row_height = max(item["height"] for item in row)
            cards_area_height += row_height
        cards_area_height += row_gap * (len(rows) - 1)

        full_width = canvas_width - (outer_padding * 2)
        full_dimensions = dict(dimensions)
        full_dimensions["card_width"] = full_width
        os_card = _compute_export_card_layout(os_row["section"], full_dimensions, metrics)

        logo_path = get_logo_path()
        has_logo = bool(logo_path)
        logo_size = max(48, int(84 * scale)) if has_logo else 0
        header_height = max(scan_metrics.height(), logo_size if has_logo else 0)

        required_height = (
            outer_padding
            + header_height
            + max(8, int(14 * scale))
            + cards_area_height
            + row_gap
            + os_card["height"]
            + outer_padding
        )

        if required_height <= canvas_height or scale <= 0.65:
            break
        scale -= 0.05

    pixmap = QPixmap(canvas_width, canvas_height)
    pixmap.fill(QColor("#000000"))

    painter = QPainter(pixmap)
    try:
        painter.setRenderHint(QPainter.TextAntialiasing, True)
        painter.setRenderHint(QPainter.Antialiasing, True)

        y = outer_padding

        logo_path = get_logo_path()
        logo_size = max(48, int(84 * scale)) if logo_path else 0
        if logo_path:
            logo_pixmap = QPixmap(logo_path)
            if not logo_pixmap.isNull():
                logo_pixmap = logo_pixmap.scaled(
                    logo_size,
                    logo_size,
                    Qt.KeepAspectRatio,
                    Qt.SmoothTransformation,
                )
                painter.drawPixmap(outer_padding, y, logo_pixmap)

        painter.setPen(QPen(QColor("#ffffff")))
        painter.setFont(scan_font)
        header_text_x = outer_padding + (logo_size + 14 if logo_size else 0)
        painter.drawText(header_text_x, y + scan_metrics.ascent(), "Flex Card")
        painter.drawText(header_text_x, y + scan_metrics.height() + scan_metrics.ascent(), f"Last scan: {specs.get('Scanned At', '-')}")

        header_height = max(scan_metrics.height() * 2, logo_size if logo_size else 0)
        y += header_height + max(8, int(14 * scale))

        def draw_card(card, x, card_y, width, height):
            painter.setPen(QPen(QColor("#f1f1f1"), 2))
            painter.setBrush(Qt.NoBrush)
            painter.drawRoundedRect(x, card_y, width, height, 24, 24)

            icon_path = card["section"].get("icon")
            icon_draw_size = icon_size if icon_path else 0
            text_x = x + card_padding + icon_draw_size + (icon_gap if icon_path else 0)
            text_y = card_y + card_padding

            if icon_path:
                icon_y = card_y + card_padding
                _draw_icon(painter, icon_path, x + card_padding, icon_y, icon_size)

            painter.setPen(QPen(QColor("#ffffff")))
            painter.setFont(title_font)
            for line in card["title_lines"]:
                painter.drawText(text_x, text_y + title_metrics.ascent(), line)
                text_y += title_metrics.height()

            text_y += max(4, int(6 * scale))
            painter.setPen(QPen(QColor("#f2f2f2")))
            painter.setFont(subtitle_font)
            for line in card["subtitle_lines"]:
                painter.drawText(text_x, text_y + subtitle_metrics.ascent(), line)
                text_y += subtitle_metrics.height()

            text_y = card_y + card_padding + card["header_height"] + max(8, int(12 * scale))
            painter.setPen(QPen(QColor("#ffffff")))
            painter.setFont(bullet_font)
            for line in card["bullet_lines"]:
                painter.drawText(x + card_padding, text_y + bullet_metrics.ascent(), line)
                text_y += bullet_metrics.height()

        for row in rows:
            row_height = max(item["height"] for item in row)
            left_x = outer_padding
            right_x = outer_padding + card_width + column_gap

            draw_card(row[0], left_x, y, card_width, row_height)
            if len(row) > 1:
                draw_card(row[1], right_x, y, card_width, row_height)

            y += row_height + row_gap

        full_width = canvas_width - (outer_padding * 2)
        draw_card(os_card, outer_padding, y, full_width, os_card["height"])
    finally:
        painter.end()
    return pixmap.save(output_path, "PNG")


# ------------------------------
# Main application window (PyQt5 GUI)
# ------------------------------
class HardwareInfoWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Flex Card")
        self.setMinimumSize(900, 650)
        self.current_specs = {}
        self.section_cards = {}
        logo_path = get_logo_path()
        if logo_path:
            self.setWindowIcon(QIcon(logo_path))
        self._apply_modern_theme()
        self._build_ui()
        self.refresh_data()

    # ------------------------------
    # Modern dark UI theme
    # ------------------------------
    def _apply_modern_theme(self):
        self.setStyleSheet(
            """
            QMainWindow, QWidget {
                background-color: #000000;
                color: #ffffff;
                font-size: 14px;
                font-family: 'Ubuntu';
            }

            QScrollArea {
                border: none;
                background-color: #000000;
            }

            QGroupBox {
                border: 2px solid #f1f1f1;
                border-radius: 20px;
                margin-top: 0px;
                padding: 16px;
                background-color: #000000;
                color: #ffffff;
            }

            QGroupBox::title {
                subcontrol-origin: margin;
                left: -9999px;
                padding: 0px;
            }

            QLabel {
                color: #ffffff;
            }

            QPushButton {
                background-color: #000000;
                color: #ffffff;
                border: 2px solid #f1f1f1;
                border-radius: 12px;
                padding: 10px 16px;
                font-weight: 600;
            }

            QPushButton:hover {
                background-color: #101010;
                border: 2px solid #ffffff;
            }

            QPushButton:pressed {
                background-color: #1a1a1a;
            }
            """
        )

    def _build_ui(self):
        main_widget = QWidget()
        root_layout = QVBoxLayout(main_widget)
        root_layout.setContentsMargins(20, 20, 20, 20)
        root_layout.setSpacing(14)

        brand_row = QHBoxLayout()
        brand_row.setSpacing(10)

        logo_path = get_logo_path()
        if logo_path:
            logo_label = QLabel()
            logo_pixmap = QPixmap(logo_path)
            if not logo_pixmap.isNull():
                logo_label.setPixmap(
                    logo_pixmap.scaled(36, 36, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                )
                brand_row.addWidget(logo_label)

        app_name_label = QLabel("Flex Card")
        app_name_label.setStyleSheet("font-size: 22px; font-weight: 700; color: #ffffff;")
        brand_row.addWidget(app_name_label)
        brand_row.addStretch(1)
        root_layout.addLayout(brand_row)

        self.scan_time_label = QLabel("Last scan: -")
        self.scan_time_label.setStyleSheet("color: #d9d9d9; font-size: 12px;")
        root_layout.addWidget(self.scan_time_label)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_content = QWidget()
        self.grid = QGridLayout(self.scroll_content)
        self.grid.setSpacing(16)

        self.cpu_group = self._make_info_group("CPU")
        self.gpu_group = self._make_info_group("GPU")
        self.ram_group = self._make_info_group("RAM")
        self.storage_group = self._make_info_group("Storage")
        self.board_group = self._make_info_group("Motherboard")
        self.bios_group = self._make_info_group("BIOS")
        self.os_group = self._make_info_group("Operating System")

        self.grid.addWidget(self.cpu_group, 0, 0)
        self.grid.addWidget(self.gpu_group, 0, 1)
        self.grid.addWidget(self.ram_group, 1, 0)
        self.grid.addWidget(self.storage_group, 1, 1)
        self.grid.addWidget(self.board_group, 2, 0)
        self.grid.addWidget(self.bios_group, 2, 1)
        self.grid.addWidget(self.os_group, 3, 0, 1, 2)

        self.scroll_area.setWidget(self.scroll_content)
        root_layout.addWidget(self.scroll_area, 1)

        button_row = QHBoxLayout()
        button_row.addStretch(1)

        self.refresh_button = QPushButton("Refresh")
        self.refresh_button.clicked.connect(self.refresh_data)
        button_row.addWidget(self.refresh_button)

        self.export_button = QPushButton("Export to .png")
        self.export_button.clicked.connect(self.export_specs)
        button_row.addWidget(self.export_button)

        root_layout.addLayout(button_row)

        self.setCentralWidget(main_widget)

    # ------------------------------
    # Section icon helpers
    # ------------------------------
    def _get_section_icon_path(self, title):
        os_name = ""
        if title == "Operating System" and self.current_specs:
            os_name = self.current_specs.get("OS", {}).get("Name", "")
        return get_section_icon_path(title, os_name=os_name)

    def _build_icon_label(self, icon_path, size=72):
        extension = os.path.splitext(icon_path)[1].lower()
        pixmap = QPixmap(size, size)
        pixmap.fill(Qt.transparent)

        if extension == ".svg":
            renderer = QSvgRenderer(icon_path)
            if not renderer.isValid():
                return None
            painter = QPainter(pixmap)
            renderer.render(painter)
            painter.end()
        else:
            source = QPixmap(icon_path)
            if source.isNull():
                return None
            pixmap = source.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation)

        icon_label = QLabel()
        icon_label.setPixmap(pixmap)
        icon_label.setFixedSize(size, size)
        return icon_label

    def _make_info_group(self, title):
        group = QGroupBox("")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(14)

        header_layout = QHBoxLayout()
        header_layout.setSpacing(16)

        icon_path = self._get_section_icon_path(title)
        icon_label = None
        if icon_path:
            icon_label = self._build_icon_label(icon_path)
        if icon_label:
            header_layout.addWidget(icon_label)

        heading_block = QVBoxLayout()
        heading_block.setSpacing(2)

        header_title = QLabel(title)
        header_title.setStyleSheet("font-size: 20px; font-weight: 500; color: #ffffff;")
        heading_block.addWidget(header_title)

        subtitle_label = QLabel("Loading...")
        subtitle_label.setStyleSheet("font-size: 12px; color: #f0f0f0;")
        subtitle_label.setWordWrap(True)
        heading_block.addWidget(subtitle_label)

        header_layout.addLayout(heading_block)
        header_layout.addStretch(1)
        layout.addLayout(header_layout)

        bullets_label = QLabel("• Loading details...")
        bullets_label.setWordWrap(True)
        bullets_label.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        bullets_label.setStyleSheet("font-size: 13px; line-height: 1.6;")
        layout.addWidget(bullets_label)

        self.section_cards[title] = {
            "subtitle": subtitle_label,
            "bullets": bullets_label,
        }
        return group

    def _dict_to_lines(self, data):
        return "\n".join(f"{k}: {v}" for k, v in data.items())

    def _gpu_to_lines(self, gpus):
        chunks = []
        for i, gpu in enumerate(gpus, start=1):
            chunks.append(
                f"GPU {i}\n"
                f"Name: {gpu.get('Name', 'N/A')}\n"
                f"Memory: {gpu.get('Memory', 'N/A')}"
            )
        return "\n\n".join(chunks)

    def _storage_to_lines(self, drives):
        chunks = []
        for drive in drives:
            chunks.append(
                f"Drive: {drive.get('Drive', 'N/A')}\n"
                f"Mount: {drive.get('Mount', 'N/A')}\n"
                f"File System: {drive.get('File System', 'N/A')}\n"
                f"Total: {drive.get('Total', 'N/A')}\n"
                f"Free: {drive.get('Free', 'N/A')}"
            )
        return "\n\n".join(chunks)

    def refresh_data(self):
        try:
            self.current_specs = collect_system_specs()
            self._update_cpu_card(self.current_specs["CPU"])
            self._update_gpu_card(self.current_specs["GPU"])
            self._update_ram_card(self.current_specs["RAM"])
            self._update_storage_card(self.current_specs["Storage"])
            self._update_motherboard_card(self.current_specs["Motherboard"])
            self._update_bios_card(self.current_specs["BIOS"])
            self._update_os_card(self.current_specs["OS"])
            self.scan_time_label.setText(f"Last scan: {self.current_specs.get('Scanned At', '-')}")
        except Exception as exc:
            QMessageBox.critical(self, "Error", f"Failed to gather system information.\n\n{exc}")

    def _set_card_content(self, section, subtitle, bullet_lines):
        card = self.section_cards.get(section)
        if not card:
            return
        card["subtitle"].setText(subtitle)
        card["bullets"].setText("\n".join(f"• {line}" for line in bullet_lines if line))

    def _update_cpu_card(self, cpu):
        name = cpu.get("Name", "N/A")
        cores = cpu.get("Physical Cores", "N/A")
        subtitle = f"{name}\n{cores}-core processor" if name != "N/A" else "CPU details"
        lines = [
            f"Logical threads: {cpu.get('Logical Threads', 'N/A')}",
            f"Max Frequency: {cpu.get('Max Frequency', 'N/A')}",
        ]
        self._set_card_content("CPU", subtitle, lines)

    def _update_gpu_card(self, gpus):
        first_gpu = _pick_preferred_gpu(gpus)
        subtitle = first_gpu.get("Name", "N/A")
        lines = [
            f"Memory: {first_gpu.get('Memory', 'N/A')}",
            f"Type: {first_gpu.get('Type', 'Unknown')}",
            f"Driver: {first_gpu.get('Driver', 'N/A')}",
        ]
        self._set_card_content("GPU", subtitle, lines)

    def _update_ram_card(self, ram):
        subtitle = f"Installed: {ram.get('Installed', ram.get('Total', 'N/A'))}"
        lines = [
            f"Usable: {ram.get('Usable', 'N/A')}",
            f"Modules: {ram.get('Module Layout', ram.get('Modules', 'N/A'))}",
            f"Speed: {ram.get('Speed', 'N/A')}",
        ]
        self._set_card_content("RAM", subtitle, lines)

    def _update_storage_card(self, drives):
        first_drive = drives[0] if drives else {"Installed": "N/A", "Usable": "N/A", "Used": "N/A", "File System": "N/A"}
        subtitle = f"Installed: {first_drive.get('Installed', 'N/A')}"
        lines = [
            f"Usable: {first_drive.get('Usable', first_drive.get('Total', 'N/A'))}",
            f"Used: {first_drive.get('Used', 'N/A')}",
            f"File System: {first_drive.get('File System', 'N/A')}",
        ]
        self._set_card_content("Storage", subtitle, lines)

    def _update_motherboard_card(self, board):
        subtitle = board.get("Model", "N/A")
        lines = [
            f"Manufacturer: {board.get('Manufacturer', 'N/A')}",
        ]
        self._set_card_content("Motherboard", subtitle, lines)

    def _update_bios_card(self, bios):
        subtitle = bios.get("BIOS Version", "N/A")
        lines = [
            f"Release Date: {bios.get('Release Date', 'N/A')}",
        ]
        self._set_card_content("BIOS", subtitle, lines)

    def _update_os_card(self, os_info):
        subtitle = os_info.get("Name", "N/A")
        lines = [
            f"Version: {os_info.get('Version', 'N/A')}",
            f"Build: {os_info.get('Build', 'N/A')}",
            f"Architecture: {os_info.get('Architecture', 'N/A')}",
        ]
        self._set_card_content("Operating System", subtitle, lines)

    def export_specs(self):
        if not self.current_specs:
            QMessageBox.warning(self, "No Data", "No system data available to export.")
            return

        default_name = f"system_specs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        dialog = QFileDialog(self, "Export Hardware Specs")
        dialog.setAcceptMode(QFileDialog.AcceptSave)
        dialog.setNameFilter("PNG Files (*.png)")
        dialog.setDefaultSuffix("png")
        dialog.setDirectory(os.path.expanduser("~"))
        dialog.selectFile(default_name)
        dialog.setOption(QFileDialog.DontUseNativeDialog, False)

        if not dialog.exec_():
            return

        selected_files = dialog.selectedFiles()
        path = selected_files[0] if selected_files else ""

        if not path:
            return

        if not path.lower().endswith(".png"):
            path += ".png"

        try:
            saved = export_specs_to_png(self.current_specs, path)
            if not saved:
                raise RuntimeError("Failed to write PNG file.")
            QMessageBox.information(self, "Export Complete", f"Specs exported to PNG:\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "Export Failed", f"Could not export file.\n\n{exc}")


# ------------------------------
# Application entry point (PyInstaller-friendly)
# ------------------------------
def main():
    app = QApplication(sys.argv)
    app.setFont(QFont("Ubuntu", 12))
    app.setStyle("Fusion")
    window = HardwareInfoWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
