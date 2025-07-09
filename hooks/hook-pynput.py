# hook-pynput.py
from PyInstaller.utils.hooks import collect_submodules, collect_data_files, collect_dynamic_libs

# Explicitly collect all submodules for pynput
# This is often more robust than just listing them in hiddenimports in the .spec file
hiddenimports = collect_submodules('pynput')

# Also explicitly collect the keyboard top-level module, as pynput is a dependency of it.
hiddenimports.append('keyboard')

# Collect platform-specific dynamic libraries if pynput relies on them (e.g., C DLLs)
# This can be crucial for the keyboard/pynput library to function correctly
binaries = collect_dynamic_libs('pynput')

# For data files: pynput usually doesn't have standalone data files.
# If it did, you would collect them here, e.g., datas = collect_data_files('pynput')
datas = [] # No known data files for pynput itself