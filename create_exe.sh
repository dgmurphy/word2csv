#!/bin/bash
# We create a virtual environment and install one package: PyInstaller.
# That package is used to create the executable.

# If necessary, create the virtual environment.
if ! [[ -d install_venv ]]; then
	python -m venv install_venv
fi
	
# Activate the environment.
if [[ "$OSTYPE" == "linux-gnu"* ]] || [[ "$OSTYPE" == "darwin"* ]]; then
	source install_venv/bin/activate
else
	source install_venv/Scripts/activate
fi
	
# This will process instantly if the packages are already installed.
pip install pyinstaller
pip install -r requirements.txt

# Create the executable.
if ! [[ -d dist ]]; then
	mkdir dist
fi
pyinstaller --onefile --hidden-import docx2txt --distpath dist src/word-2-excel.py

# Deactivate the virtual environment.
deactivate
