# Python Scripts

A collection of Python scripts for automation, data processing, and general utilities.
Created for learning and demonstration purposes.

# BOM Comparison Tool

## Description
The BOM Comparison Tool is a PyQt5-based GUI application designed to compare two Bill of Materials (BOM) Excel files. It allows users to select two BOM files, apply various filters, and generate a comparison report in Excel format.

## Features
- Select two BOM files for comparison.
- Apply filters to remove zero quantity items, new zero quantity items, deprecated zero quantity items, and unchanged items.
- Generate a comparison report in Excel format.
- Write the RefDes list from the merged DataFrame to a text file.

## Classes and Functions
### Classes
- `BOMComparisonApp(QMainWindow)`: Main application window for the BOM comparison tool.

### Functions
- `__init__()`: Initializes the main application window.
- `init_ui()`: Sets up the user interface components.
- `create_button(text, callback)`: Creates a QPushButton with a specified callback function.
- `select_file(label, file_attr)`: Generic method to select a file and update the corresponding label.
- `select_file1()`: Opens a file dialog to select the first BOM file.
- `select_file2()`: Opens a file dialog to select the second BOM file.
- `run_comparison()`: Executes the BOM comparison process and generates the output files.
- `get_output_file_name()`: Generates the output file name based on the selected BOM files.
- `read_excel_files()`: Reads the selected BOM Excel files into pandas DataFrames.
- `merge_and_filter_dataframes(df1, df2)`: Merges and applies filters to the BOM DataFrames.
- `handle_existing_file(output_file)`: Handles cases where the output file already exists.
- `write_to_excel(output_file, df1, df2, df_merged)`: Writes the comparison results to an Excel file.
- `handle_permission_error()`: Handles permission errors when writing to a file.
- `modify_excel_file(output_file)`: Modifies the Excel file using openpyxl to add formatting and filters.
- `write_refdes_to_txt(df_merged)`: Writes the RefDes list from the merged DataFrame to a text file.

## Constants
- `COLUMNS`: List of column names to be used when reading the BOM Excel files.
- `OUTPUT_DIR`: Directory where the output files will be saved.
- `EXCLUDED_PREFIX`: Prefix used to exclude certain files.

## Usage
1. Run the script to launch the GUI application.
2. Use the interface to select two BOM files.
3. Apply filters as needed.
4. Click "Run Comparison" to generate the comparison report.

---

# PDF Highlighter Tool

## Description
The PDF Highlighter Tool is a PyQt5-based GUI application that allows users to highlight search terms in PDF files. Users can select one or two PDF files, a text file containing search terms, and a highlight color. The tool processes the PDFs and saves the highlighted versions.

## Features
- Select one or two PDF files for highlighting.
- Select a text file containing search terms.
- Choose a highlight color from a predefined list.
- Generate highlighted PDFs and save them to the output directory.

## Classes and Functions
### Classes
- `PDFHighlighterWorker(QThread)`: Worker thread for processing PDF files.
- `PDFHighlighterApp(QMainWindow)`: Main application window for the PDF highlighter tool.

### Functions
- `__init__()`: Initializes the main application window.
- `init_ui()`: Sets up the user interface components.
- `update_color_preview(color)`: Updates the color preview box based on the selected color.
- `select_pdf1()`: Opens a file dialog to select the first PDF file.
- `select_pdf2()`: Opens a file dialog to select the second PDF file.
- `select_txt()`: Opens a file dialog to select the text file containing search terms.
- `select_color(color)`: Sets the highlight color.
- `handle_existing_file(output_file)`: Handles cases where the output file already exists.
- `start_processing()`: Starts the PDF highlighting process.
- `update_progress(value)`: Updates the progress bar.
- `log_message(message)`: Logs messages to the QTextEdit.
- `processing_finished(message)`: Handles the completion of the PDF processing.

## Constants
- `COLORS`: Dictionary of predefined highlight colors.
- `DefaultColor`: Default highlight color.
- `DefaultColorRGB`: RGB values of the default highlight color.
- `OUTPUT_DIR`: Directory where the output files will be saved.

## Usage
1. Run the script to launch the GUI application.
2. Use the interface to select one or two PDF files.
3. Select a text file containing search terms.
4. Choose a highlight color.
5. Click "Highlight PDFs" to generate the highlighted PDFs.

---

# BLF File Processor

## Description
The BLF File Processor is a script that processes BLF files to extract and analyze CAN bus data. It reads BLF files, filters the data, calculates durations, adds off time periods, and generates a cycle count. The results are saved as Excel files and images.

## Features
- List all BLF files in the current directory.
- Read and filter CAN bus data from BLF files.
- Calculate durations of each state.
- Add off time periods to the results.
- Generate a cycle count for specific states.
- Save the results as Excel files and images.

## Functions
- `list_blf_files()`: Lists all BLF files in the current directory.
- `read_and_filter(df)`: Reads and filters the DataFrame based on specific columns.
- `add_time_columns(df)`: Calculates the duration of each state.
- `add_off_time_periods(df)`: Adds off time periods to the results.
- `add_cycle_count(result_df)`: Adds a cycle count for specific states.
- `save_picture(df, file_name)`: Saves the DataFrame as an image.
- `save_to_excel(df, file_name)`: Saves the DataFrame to an Excel file.

## Usage
1. Place the BLF files in the current directory.
2. Run the script to process the BLF files.
3. The results will be saved as Excel files and images in the current directory.
