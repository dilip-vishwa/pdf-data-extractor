
def convert_excel_to_csv_by_parsing_cells_by_rows_and_columns(file_path):
    import openpyxl
    import csv

    # Assume excel_filepath and csv_filepath are passed as arguments to the function
    # For demonstration, let's assume they are available or define them here for a snippet
    # excel_filepath = "input.xlsx"
    csv_filepath = "output.csv"

    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active  # Or specify sheet by name: workbook['Sheet1']

    with open(csv_filepath, 'w', newline='', encoding='utf-8', errors='ignore') as csv_file:
        csv_writer = csv.writer(csv_file)
        for row in sheet.iter_rows():
            row_values = []
            for cell in row:
                # Check if the cell has a fill color and if it's a shade of green.
                # openpyxl uses AARRGGBB format for colors (Alpha, Red, Green, Blue).
                # Common green colors might include 'FF00FF00' (pure green), 'FF92D050' (light green), 'FF70AD47' (darker green).
                # You may need to adjust the color codes to match the specific green(s) in your Excel file.
                # This example checks for a few common green hex codes.
                is_green = False
                if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb:
                    color_rgb = str(cell.fill.start_color.rgb).upper()
                    # Check for common green hex codes. Add or modify as needed for your specific greens.
                    if color_rgb in ['FF00FF00', 'FF92D050', 'FF70AD47', 'FFC6EFCE', 'FFE2EFD9']: # C6EFCE is a common light green fill
                        is_green = True

                if is_green:
                    row_values.append("Need data here")
                else:
                    row_values.append(cell.value)

            csv_writer.writerow(row_values)


if __name__ == "__main__":
    file_path = "/home/adming/project/find_excel_data_fill_label/SPECIFICATION_SUMMARY_SHEET-BLANK.xlsx" 
    convert_excel_to_csv_by_parsing_cells_by_rows_and_columns(file_path)