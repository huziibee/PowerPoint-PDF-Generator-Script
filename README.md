# README for PowerPoint PDF Generator Script

## Overview

This Python script automates the process of creating personalized PowerPoint presentations and converting them into PDFs. It reads a list of names from a specified CSV file, uses these names to replace placeholders in PowerPoint templates, and then generates a PDF for each customized presentation.

## Video Walk-through 
[![YouTube Tutorial](https://img.youtube.com/vi/QUoLlPZyYoQ/maxresdefault.jpg)](https://youtu.be/QUoLlPZyYoQ?si=jkJZtFGAb3uyzmLZ)

## Requirements

- **Python** installed on your system.
- The following Python packages:
  - `os` (standard library)
  - `comtypes`
  - `python-pptx`
  - `pandas`

## Installation

Before running the script, ensure you have the necessary Python packages installed. You can install them using pip:

```bash
pip install comtypes python-pptx pandas
```

## Templates and CSV File

- PowerPoint templates named `"short name.pptx"` and `"long name.pptx"` are required.
- A CSV file with a column named `'Name'` containing the names for the presentations.
- **It is mandatory that the PowerPoint templates and the CSV file are located in the working directory of the script.**
- The URLs to download these templates can be found in the Slack's Canvas notes.

## Usage

1. **Prepare Your Environment**:
   - Install the required Python packages.
   - Download the PowerPoint templates from the provided URL in Slack's Canvas notes and place them in the script's working directory.
   - Ensure your CSV file is also in the script's working directory.

2. **Configure the Script**:
   - Set `file_path` to the path of your CSV file.
   - Modify `short_template_path`, `long_template_path`, and `output_folder` as needed.

3. **Run the Script**:
   - Execute the script in your Python environment.
   - The script will automatically confirm the working directory and process the names from the CSV file to create personalized presentations and convert them to PDFs.

4. **Output**:
   - Check the specified output folder for the generated PDFs.

## Notes

- The script currently processes only the first slide in each PowerPoint template.
- Names with more than one space are skipped to maintain formatting.
- The script must be run in the directory containing the necessary templates and CSV file.

## Limitations

- Designed for specific template structures; may require modifications for different templates or additional features.
- Requires the necessary files to be in the working folder for successful execution.

## Troubleshooting

For any issues or further customization, refer to the script comments or contact the script maintainer. Remember to check the Slack's Canvas notes for updates or additional resources related to this script.
