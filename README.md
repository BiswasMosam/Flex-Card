# Windows Hardware Info App (PyQt5)

A Windows desktop application in Python that collects and displays detailed hardware/system information using **WMI** and **psutil**.

## Features

- CPU: name, physical cores, logical threads, current/max frequency
- GPU: name and memory (when available)
- RAM: total, available, used, usage %, speed (when available)
- Storage: all detected drives with total/free space
- Motherboard: manufacturer and model
- BIOS: version and release date
- OS: name, version, build, architecture
- Refresh button for live re-scan
- Export button to save specs as `.txt`

## Setup

```powershell
cd D:\Garage\python\hardware_info_app
pip install -r requirements.txt
python app.py
```

## Build single EXE with PyInstaller

```powershell
cd D:\Garage\python\hardware_info_app
pyinstaller --noconfirm --onefile --windowed --name HardwareInfo app.py
```

Generated file will be in:

- `dist/HardwareInfo.exe`

## Notes

- This app is Windows-focused (uses WMI).
- Some fields (GPU memory, RAM speed) depend on system/driver support.
