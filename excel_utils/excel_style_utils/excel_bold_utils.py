from collections import defaultdict
from excel_utils.excel_line_wise import split_excel_coordinate_into_row_and_column


def get_all_bold_cells_in_excel(sheet):
    bold_cells_row_wise = defaultdict(list)
    for row in sheet.iter_rows():
        for cell in row:
            if cell.font and cell.font.bold and cell.value is not None:
                bold_cells_row_wise[cell.row].append(cell)

    bold_cells = {}
    for row, cells in bold_cells_row_wise.items():
        bold_cells[row] = " ".join([str(b.value) for b in cells])

    return bold_cells

def get_immediate_lowest_number_key_from_dict_with_number_as_key(dict, current_row):
    keys = list(dict.keys())
    keys.sort(reverse=True)
    return [k for k in keys if k < current_row][0]


def add_bold_text_in_concerned_row_data(row_data, bold_cells):
    output = {}
    for coordinate, cell_data in row_data.items():
        row, column = split_excel_coordinate_into_row_and_column(coordinate)
        add_this_bold_text = bold_cells[get_immediate_lowest_number_key_from_dict_with_number_as_key(bold_cells, row)]
        output[coordinate] = [cell_data[0], add_this_bold_text] + cell_data[1:]
    return output
