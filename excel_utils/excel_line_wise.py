
def get_excel_data_in_line_wise_to_query_to_llm(cell_information_list):
    data = {}
    for coordinate, cell in cell_information_list.items():
        row, column = split_excel_coordinate_into_row_and_column(coordinate)
        if row not in data:
            data[row] = {}
        data[row][column] = cell
    
    return data


def split_excel_coordinate_into_row_and_column(coordinate):
    try:
        row = int(coordinate[1:])
        column = coordinate[0]
    except Exception as e:
        try:
            row = int(coordinate[2:])
            column = coordinate[0:2]
        except Exception as e:
            row = int(coordinate[3:])
            column = coordinate[0:3]

    return row, column


def construct_prompt_for_row_wise_data(output_line_wise_data):
    prompt = "You are an expert in reading data from pdf content. Give me answer in short and simple words. No need to give any explanation. If answer is expected from given options, then give answer from given options only. This data is needed to be filled in excel sheet. Give me answer in short and simple words. If you don't know the answer, then say 'NA'. Try to match with the options if given in prompt"

    for row, columns in output_line_wise_data.items():
        data = ""
        for column, cell in columns.items():
            data += f"Give data related about{" on '" + str(cell[1]) + "'" if cell[1] else ""} {" on '" + str(cell[2]) + "'" if cell[2] else ""} {" on '" + str(cell[3]) + "'" if cell[3] else ""} {". Options are '" + str(cell[5]) + "'" if len(cell) == 6 and cell[5] else ""}. Prefix with {str(cell[0]):}(answer here). \n\n"
        
        output_line_wise_data[row]['prompt'] = data
        
    return output_line_wise_data

def combine_multiple_rows_in_single_prompt(prompts_by_row, number_of_rows_in_a_prompt=10):
    prompt_list = [p['prompt'] for k, p in prompts_by_row.items()]
    batched_prompts = [prompt_list[i:i + number_of_rows_in_a_prompt] for i in range(0, len(prompt_list), number_of_rows_in_a_prompt)]
    return batched_prompts

