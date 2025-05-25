# CSV Editor - Portable Version

A portable application for filtering and updating CSV/Excel files.

## Quick Start

1. Download the latest release from the [Releases](https://github.com/yourusername/CSVEditor/releases) page
2. Extract the ZIP file
3. Double-click `run_csv_editor.bat` to start the application

## Manual Setup (Alternative)

If you prefer to set up manually:

1. Download WinPython 3.10 64-bit from: https://winpython.github.io/
2. Extract WinPython to this folder
3. Install required dependencies by running:
   ```
   .\python-3.10.11.amd64\python.exe -m pip install openpyxl
   ```
4. Double-click `run_csv_editor.bat` to start the application

## Features

- Filter CSV/Excel files based on multiple conditions
- Update values in CSV/Excel files
- Support for both CSV and Excel (.xlsx) files
- No installation required - just unzip and run!

## Requirements

- Windows operating system
- No Python installation required (included in the bundle)
- No administrator rights required

## Troubleshooting

If you encounter any issues:
1. Make sure WinPython is extracted to the same folder as the application
2. Verify that openpyxl is installed by running the pip install command above
3. Check that the Python version in run_csv_editor.bat matches your WinPython folder name

## Development

The portable package is automatically built and released when a new version tag is pushed to the repository. To create a new release:

1. Update the version number in your code
2. Create and push a new tag:
   ```bash
   git tag v1.0.0
   git push origin v1.0.0
   ```
3. The GitHub Action will automatically:
   - Download WinPython
   - Install dependencies
   - Create the portable package
   - Create a new release with the download link
