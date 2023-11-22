# README for PowerPoint PDF Generator Script

## Overview

This Python script automates the creation of personalized PowerPoint presentations and their conversion into PDF format. It dynamically reads names from a selected CSV file, applies these names to placeholders in PowerPoint templates, and generates a PDF for each customized presentation.

## Key Features

- **Automated Personalization**: Uses names from a CSV file to personalize PowerPoint templates.
- **PDF Conversion**: Converts each customized PowerPoint presentation into a PDF file.
- **Template Selection**: Chooses between short and long name templates based on name length.
- **Dynamic CSV File Selection**: Lists all CSV files in the current directory for user selection.
- **Interactive Directory Confirmation**: Allows users to confirm or change the working directory.

## Requirements

- Python with the following packages:
  - `os` (standard library)
  - `comtypes`
  - `python-pptx`
  - `pandas`

## Installation

To install the necessary Python packages, use pip:

```bash
pip install python-pptx pandas
```

**Note**: The `comtypes` package may require installation with administrative privileges. To install `comtypes`, run:

```bash
pip install comtypes
```

If you encounter permission issues, try running the command in an administrative command prompt or add the `--user` flag to install it in the user directory:

```bash
pip install comtypes --user
```

## Usage

1. **Ensure Required Files**: Place the `short name.pptx` and `long name.pptx` templates in the same directory as the script.
2. **Run the Script**: Execute `python converter.py` in your command line.
3. **Confirm or Change Directory**: When prompted, confirm the current directory or enter a new one.
4. **Select a CSV File**: Choose a CSV file from the listed options in the current directory.
5. **Check Output**: The script generates PDFs in an `output_pdfs` folder within the same directory.

## Script Workflow

- **Directory Confirmation**: The script first confirms the current working directory or changes it based on user input.
- **CSV File Selection**: Lists all CSV files in the current directory for the user to select.
- **Name Processing**: Reads names

