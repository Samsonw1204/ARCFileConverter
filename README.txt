# Excel to CSV Converter

## Overview
This program converts `.xls` Excel files into a properly formatted `.csv` file for Red Cross Learning Center class setup.

## How to Use
1. Download `ARCFileConverter.exe` from the [Releases](https://github.com/your-username/repo-name/releases) page.
2. Double-click the `.exe` to start the program.
3. Follow the on-screen prompts to convert your file.

## Requirements
- Java Runtime Environment (if not bundled)

## Notes
- The program only supports `.xls` files (not `.xlsx`).
- If errors occur during processing, skipped rows will be logged in `skipped_rows.log`.
- Ensure that the input file is an export from RecDesk; it will only work for files formatted in that way.

## Troubleshooting
- **Java Not Recognized**:
  - Ensure Java is installed and added to your system PATH.
  - On Windows, re-run the Java installer and ensure "Add Java to PATH" is selected.
  - If none of that works, I reccomend you ask ChatGPT for help; it is very good at troubleshooting these things.
- **Error: Invalid File Path**:
  - Make sure the path you entered points to a valid `.xls` file.
  - If copying the path from File Explorer, ensure there are no extra quotes or spaces.

## Contact
For issues or questions, contact the program developer: Samson B. Wheelock at Samsonw1204@gmail.com
