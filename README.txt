# Excel to CSV Converter

## Overview
This program converts `.xls` Excel files into a properly formatted `.csv` file for Red Cross Learning Center class setup.

## How to Use
1. Download `ARCFileConverter_Releases.zip` directory from the [Releases](https://github.com/Samsonw1204/ARCFileConverter/releases) page.
2. Extract the project files from the .zip file.
3. Double-click the `.exe` to start the program.
4. Follow the on-screen prompts to convert your file.

## Requirements
- Java Runtime Environment (Bundled into program software)

## Notes
- The program only supports `.xls` files (not `.xlsx`).
- If errors occur during processing, skipped rows will be logged in `skipped_rows.log`.
- Ensure that the input file is an export from RecDesk; it will only work for files formatted in that way.

## Security
- The .exe file provided to run this program does not include a security certificate from a trusted authority. Purchasing such a certificate was not feasible due to their cost.
- Why this happens: 
    - Windows Defender and other antivirus programs are designed to flag unsigned executables as a precaution against malicious software. 
    - This is not an indication of actual harm but rather a lack of verification by a recognized authority.
- Why this program is safe:
    - This program was built entirely using Java, a widely used programming language.
    - The executable (.exe file) was generated directly from the program’s source code, which you can inspect in this GitHub repository to verify its contents.
    - The included files, such as the jre/ folder, are standard components from the official Java distribution.
- Running the Program:
    - Double click the "ARCFileConverter.exe" application in the ARCFileConverter directory. 
    - Windows Defender will likely prohibit the application from running due to the lack of a security certificate (see above); however, you can bypass this by clicking "More Info" and then "Run Anyway."
- If you have any doubts or concerns, you can clone the repository and rebuild the program using the provided source code to ensure its integrity.

## Troubleshooting
- **Java Not Recognized**:
  - Ensure Java is installed and added to your system PATH.
  - On Windows, re-run the Java installer and ensure "Add Java to PATH" is selected.
  - If none of that works, I recommend you ask ChatGPT for help; it is very good at troubleshooting these things.
- **Error: Invalid File Path**:
  - Make sure the path you entered points to a valid `.xls` file.
  - If copying the path from File Explorer, ensure there are no extra quotes or spaces.

## Contact
For issues or questions, contact the program developer: Samson B. Wheelock at Samsonw1204@gmail.com
