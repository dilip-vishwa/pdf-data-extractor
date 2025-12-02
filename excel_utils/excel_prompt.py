from excel_utils.excel_style_utils.excel_bold_utils import get_all_bold_cells_in_excel
from excel_utils.excel_style_utils.excel_bold_utils import add_bold_text_in_concerned_row_data
from excel_utils.excel_line_wise import split_excel_coordinate_into_row_and_column
from excel_utils.excel_line_wise import combine_multiple_rows_in_single_prompt
from excel_utils.excel_line_wise import construct_prompt_for_row_wise_data
from excel_utils.excel_line_wise import get_excel_data_in_line_wise_to_query_to_llm
from excel_utils.excel_validations import get_validation_map
from excel_utils.test_data import check_if_output_is_correct
from excel_utils.excel_adjacent_cell_utils import find_label_of_green_cell_using_rules
from excel_utils.excel_open_utils import get_excel_sheet
from excel_utils.excel_style_utils.excel_color_utils import get_colored_cells
from collections import defaultdict


def get_prompt_from_excel(file_path):
    # find_green_cell_labels(file_path)
    excel_workbook, sheet = get_excel_sheet(file_path)
    green_cells = get_colored_cells(sheet)
    bold_cells = get_all_bold_cells_in_excel(sheet)
    print(green_cells)
    outputs = {}
    for green_cell in green_cells:
        max_parsing_times = {
            'left': 4,
            'top': 1,
            'right': 1,
            'bottom': 4
        }

        output = find_label_of_green_cell_using_rules(green_cell, sheet, max_parsing_times)
        print(output)
        outputs[green_cell.coordinate] = output

    check_if_output_is_correct(outputs)

    validation_map = get_validation_map(excel_workbook, sheet)
    print(validation_map)
    for coordinate, validation in validation_map.items():
        if coordinate in outputs:
            outputs[coordinate].append(validation)

    outputs = add_bold_text_in_concerned_row_data(outputs, bold_cells)
    lines_for_llm = get_excel_data_in_line_wise_to_query_to_llm(outputs)
    prompt = construct_prompt_for_row_wise_data(lines_for_llm)
    combined_prompt = combine_multiple_rows_in_single_prompt(prompt, number_of_rows_in_a_prompt=4)

    # all_llm_response = []
    # for cp in combined_prompt:
    #     all_llm_response.append(send_prompt_to_llm("\n\n".join(cp)))
    # print(all_llm_response)
    print("Done")
    return combined_prompt


if __name__ == "__main__":
    file_path = "/home/adming/project/pdf-data-extractor/dataset/SPECIFICATION_SUMMARY_SHEET-BLANK.xlsx"
    get_prompt_from_excel(file_path)
