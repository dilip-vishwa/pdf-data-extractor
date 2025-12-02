
def get_colored_cells(sheet, target_color_hex="FFE2EFD9"):
    green_cells = []
    
    for row in sheet.iter_rows():
        for cell in row:
            if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb == target_color_hex:
                green_cells.append(cell)
                
    return green_cells

def check_if_cell_is_colored(cell, target_color_hex="FFE2EFD9"):
    if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb == target_color_hex:
        return True
    return False

def check_if_cell_has_no_content_and_no_color_hence_empty(cell, sheet=None):
    if isinstance(cell, str):
        cell = sheet[cell]
    return cell.value is None and cell.fill and cell.fill.start_color and cell.fill.start_color.rgb == "00000000"

def check_if_given_cell_is_text_or_green_or_empty(cell, sheet=None):
    if isinstance(cell, str):
        cell = sheet[cell]
    if cell is None:
        return 'empty'
    if cell.value is not None:
        return 'text'
    if check_if_cell_is_colored(cell):
        return 'green'
    if check_if_cell_has_no_content_and_no_color_hence_empty(cell, sheet):
        return 'empty'

    return 'unknown'