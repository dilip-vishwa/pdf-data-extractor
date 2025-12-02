
from excel_utils.excel_line_wise import split_excel_coordinate_into_row_and_column
import openpyxl

def get_excel_coordinates_and_answer_from_llm_answer(answers):
    final_data = []
    for answer in answers.split('\n'):
        answer = answer.strip()
        if not answer:
            continue
        if '(' in answer and ')' == answer[-1]:
            # split_excel_coordinate_into_row_and_column(answer.split('('))
            final_data.append((answer.split('(')[0], " ".join(answer.split('(')[1:])[:-1]))

        elif ':' in answer and ')' != answer[-1]:
            # split_excel_coordinate_into_row_and_column(answer.split(':'))
            final_data.append((answer.split(':')[0], " ".join(answer.split(':')[1:])))

        elif ':' not in answer and ')' != answer[-1]:
            # split_excel_coordinate_into_row_and_column(answer.split(' '))
            final_data.append((answer.split(' ')[0], " ".join(answer.split(' ')[1:])))
        else:
            raise Exception(f"Invalid answer format: {answer}")

    return final_data


def get_answer_from_llm_response(llm_responses):
    all_response_with_excel_coordinates = []
    for llm_response in llm_responses:
        all_response_with_excel_coordinates.extend(get_excel_coordinates_and_answer_from_llm_answer(llm_response['answer']))

    return all_response_with_excel_coordinates


def fill_excel(file_path, data_map, output_path):
    """
    Fills the Excel file with data at specified coordinates.
    
    Args:
        file_path (str): Path to the source Excel file.
        data_map (dict): Dictionary mapping coordinates (e.g., 'D4') to values.
        output_path (str): Path to save the modified Excel file.
    """
    try:
        # import shutil
        # shutil.copy2(file_path, output_path)

        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        
        for coord, value in data_map:
            if value is not None:
                sheet[coord] = value
                
        wb.save(output_path)
        print(f"Successfully saved to {output_path}")
        
    except Exception as e:
        print(f"Error writing to Excel {output_path}: {e}")
