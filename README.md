# Excel to CSV Documentation

## 1. Overview

Provide a brief description of what the program does and how it solves a problem or meets a need. This helps readers understand the context right away.

**Example:**  
This program converts Excel files into CSV format, a simpler and more accessible format for data manipulation and analysis. Users can easily drag and drop an Excel file into the app, and the program will handle the conversion of all sheets into individual CSV files, simplifying the process for users who need to work with large datasets.

---

## 2. Features

Highlight key features of the program. This section is important to showcase the functionality the program provides.

**Example:**
- Drag-and-drop interface for ease of use.
- Converts multiple sheets from an Excel file into individual CSV files.
- Displays progress in a log output window.
- Option to clear log output for fresh updates.
- Automatically saves the CSV files in the same directory as the original Excel file.

---

## 3. Requirements/Dependencies

List any libraries, software, or system requirements needed to run the program.

**Example:**
- Python 3.x
- Pandas (for handling Excel and CSV files)
- Tkinter (for GUI)
- tkinterdnd2 (for drag-and-drop functionality)

---

## 4. Usage Instructions

Provide a more detailed "how-to" for users. Step-by-step instructions on how to use the program.

**Example:**
1. Launch the application.
2. Drag and drop an Excel file (.xlsx) onto the designated area.
3. The program will process each sheet in the Excel file and save it as a separate CSV file in the same directory.
4. Monitor the log output for updates on the conversion process.
5. Use the “Clear” button to reset the log output window.

---

## 5. Known Issues & Limitations

This section can mention any known bugs or limitations, ensuring users are aware of them.

**Example:**
- Does not support older Excel file formats (.xls).
- Error handling is basic; files without any sheets or with special characters may cause issues.
- The program does not support Excel files that contain complex formulas or formatting.

---

## 6. Future Enhancements

This section could discuss any potential improvements that could be made in the future.

**Example:**
- Support for older Excel file formats (.xls).
- Option to choose the destination folder for CSV files.
- Multi-threading to handle large files faster.
- More robust error handling for edge cases (empty sheets, special characters, etc.).

---

## 7. Testing

The program was tested using basic Excel sheets composed of multiple worksheets. The tests confirmed that the program can:
- Correctly process Excel files with multiple sheets.
- Convert each sheet to an individual CSV file.
- Save the CSV files in the same folder as the original Excel file.

Testing was conducted with sample Excel files containing varying numbers of sheets to ensure the program handles typical use cases.
