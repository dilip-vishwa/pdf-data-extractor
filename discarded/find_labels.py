# import openpyxl
# from openpyxl.utils import get_column_letter
# import re

# def get_merged_cell_value(sheet, cell):
#     """
#     If a cell is part of a merged range, return the value of the top-left cell of that range.
#     Otherwise, return the cell's own value.
#     Returns the value and the coordinate of the top-left cell.
#     """
#     for merged_range in sheet.merged_cells.ranges:
#         if cell.coordinate in merged_range:
#             # Get the top-left cell of the merged range
#             top_left_cell = sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
#             return top_left_cell.value, top_left_cell
#     return cell.value, cell

# def is_bold(cell):
#     return cell.font and cell.font.bold

# def has_numbering(text):
#     if not text or not isinstance(text, str):
#         return False
#     # Matches "1.", "1)", "A.", "A)", "(1)", etc.
#     return bool(re.match(r'^(\d+|[A-Za-z])[\.\)]', text.strip())) or bool(re.match(r'^\(\d+\)', text.strip()))

# def get_score(cell, position_score):
#     """
#     Calculate a score for a potential label cell.
#     """
#     score = 0
#     value = cell.value
    
#     if value is None:
#         return -1 # Empty cells are not labels

#     # Convert to string for regex check
#     str_value = str(value)
    
#     # 1. Formatting: Bold
#     if is_bold(cell):
#         score += 3
        
#     # 2. Content: Numbering
#     if has_numbering(str_value):
#         score += 2
        
#     # 3. Position preference
#     score += position_score
    
#     return score

# def find_green_cell_labels(file_path, target_color_hex="FFE2EFD9"):
#     try:
#         wb = openpyxl.load_workbook(file_path, data_only=True)
#         # Assuming we want to check all sheets, or just the active one? 
#         # User didn't specify, but usually active sheet is the main one.
#         sheet = wb.active
#         print(f"Processing sheet: {sheet.title}")
        
#         green_cells = []
        
#         for row in sheet.iter_rows():
#             for cell in row:
#                 if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb == target_color_hex:
#                     green_cells.append(cell)
                    
#         print(f"Found {len(green_cells)} green cells.")
        
#         results = []
        
#         for cell in green_cells:
#             row = cell.row
#             col = cell.column
            
#             # Define neighbors: (row_offset, col_offset, position_name, position_score)
#             # Preference: Left (high), Top (high), Right (low), Bottom (low)
#             neighbors_def = [
#                 (0, -1, "Left", 1.5),
#                 (-1, 0, "Top", 1.5),
#                 (0, 1, "Right", 0.5),
#                 (1, 0, "Bottom", 0.5)
#             ]
            
#             candidates = []
            
#             for r_off, c_off, pos_name, pos_score in neighbors_def:
#                 n_row = row + r_off
#                 n_col = col + c_off
                
#                 if n_row > 0 and n_col > 0:
#                     try:
#                         neighbor_cell = sheet.cell(row=n_row, column=n_col)
#                         # Check if neighbor is merged
#                         val, actual_cell = get_merged_cell_value(sheet, neighbor_cell)
                        
#                         if val is not None:
#                             score = get_score(actual_cell, pos_score)
#                             if score > 0:
#                                 candidates.append({
#                                     "label": val,
#                                     "position": pos_name,
#                                     "score": score,
#                                     "coordinate": actual_cell.coordinate
#                                 })
#                     except ValueError:
#                         # Cell out of bounds
#                         pass

#             # Sort candidates by score descending
#             candidates.sort(key=lambda x: x['score'], reverse=True)
            
#             best_label = candidates[0] if candidates else None
            
#             results.append({
#                 "green_cell": cell.coordinate,
#                 "label_info": best_label,
#                 "all_candidates": candidates
#             })
            
#             print(f"Green Cell {cell.coordinate}:")
#             if best_label:
#                 print(f"  Best Label: '{best_label['label']}' (Position: {best_label['position']}, Score: {best_label['score']})")
#             else:
#                 print(f"  No label found.")

#         return results

#     except Exception as e:
#         print(f"Error: {e}")
#         return []

# if __name__ == "__main__":
#     file_path = "/home/adming/project/find_excel_data_fill_label/SPECIFICATION_SUMMARY_SHEET-BLANK.xlsx"
#     find_green_cell_labels(file_path)
