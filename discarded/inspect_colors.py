# import openpyxl

# def inspect_colors(file_path):
#     try:
#         wb = openpyxl.load_workbook(file_path, data_only=True)
#         sheet = wb.active
        
#         colors = set()
        
#         print(f"Inspecting {file_path}...")
#         for row in sheet.iter_rows():
#             for cell in row:
#                 if cell.fill and cell.fill.start_color:
#                     # Capture both index and rgb if available
#                     color_info = (cell.fill.start_color.index, cell.fill.start_color.rgb, cell.fill.start_color.theme)
#                     if color_info not in colors:
#                         colors.add(color_info)
#                         print(f"Cell {cell.coordinate}: Index={cell.fill.start_color.index}, RGB={cell.fill.start_color.rgb}, Theme={cell.fill.start_color.theme}, Type={cell.fill.fill_type}")

#     except Exception as e:
#         print(f"Error: {e}")

# if __name__ == "__main__":
#     # Using the filename found in the previous list_dir
#     file_path = "/home/adming/project/find_excel_data_fill_label/SPECIFICATION_SUMMARY_SHEET-BLANK.xlsx"
#     inspect_colors(file_path)
