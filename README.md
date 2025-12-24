# MSPubConverter

A small utility to convert Microsoft Publisher (.pub) files to PDF using Microsoft Publisher's COM API.

## Key Features

- Convert all `.pub` files under a selected folder to PDF.
- Optionally preserve the input folder structure under the chosen output folder.
- Generates space-free camelCase filenames and ensures unique filenames (appends `-2`, `-3`, ... on collisions).

## Requirements

- Microsoft Publisher installed (Windows only). This tool uses COM automation and only works on Windows.
- Python 3.8 or newer.
- Python packages: `pywin32`, `tqdm`. `tkinter` is used for folder/dialog UI (usually included with Python on Windows).

Install dependencies:

```bash
pip install pywin32 tqdm
```

## Usage

Run the script from a command prompt (or double-click the file in Explorer):

```bash
python MSPubConverter.py
```

The script will open two folder-selection dialogs:

- Select the parent folder that contains the `.pub` files to convert.
- Select the output folder where PDFs should be written.

You will then be asked whether to preserve the input folder structure under the output folder. Choosing "Yes" recreates relative subfolders; "No" places all PDFs directly into the chosen output folder.

The script shows a progress bar and logs conversion activity. If Publisher fails to export a file, the script logs a warning and continues.

## Behavior notes

- Filenames: the script converts `.pub` base names to camelCase (removes non-alphanumeric chars and spaces). If the generated PDF name already exists, it appends `-2`, `-3`, etc., to make it unique.
- Output verification: after export the script checks that the file exists and logs errors if it does not.

## Troubleshooting

- "Publisher did not open the requested file" — ensure the `.pub` file is not corrupted and Microsoft Publisher can open it manually.
- COM/pywin32 errors — ensure `pywin32` is installed and you're running on Windows with Publisher installed and licensed.
- Permission errors — run the script from an account with file-system and Publisher automation permissions.

## Files

- Main script: [MSPubConverter.py](MSPubConverter.py)
- License: [LICENSE.txt](LICENSE.txt)

## License

See [LICENSE.txt](LICENSE.txt).

---
Generated from the behavior of `MSPubConverter.py`.
