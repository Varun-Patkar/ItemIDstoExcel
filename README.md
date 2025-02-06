# Item IDs to Excel

This folder contains files extracted from the Kingdom Come Deliverance mod available on Nexus Mods: https://www.nexusmods.com/kingdomcomedeliverance2/mods/87?tab=description

## Contents:

- Various item files (e.g., item.txt, item**autotests.txt, item**alchemy.txt, item\_\_aux.txt, etc.)
- Python script (process.py) to parse these files and export the item data into an Excel file.

## What It Does:

The Python script processes the mod files to extract item IDs and properties, then writes the information into an Excel file (KCD2Items.xlsx). Tables are created per file with headers for easy sorting and searching. This is especially useful for sorting equipment by attributes (such as noise) to suit stealth build requirements.

## Requirements:

- Python (version 3.x recommended)
- [XlsxWriter](https://xlsxwriter.readthedocs.io/en/latest/)

Install it using:

`pip install XlsxWriter`

## Usage:

1. Place the process.py script in the folder that contains all the item files.
2. Open a terminal, navigate to the folder, and run:

`python process.py`

3. The script will generate an Excel file named KCD2Items.xlsx with tables for each item file.

## Notes:

- This tool works with the English version of the mod files.
- It is intended for personal use to help sort and analyze item data (for example, organizing equipment by noise level for stealth builds).
- Although the script is shared here, it is not distributed as a standalone mod and is provided as-is for your convenience.

Enjoy sorting and planning your builds!

License:
This project is licensed under the MIT License. See the LICENSE file for details.
