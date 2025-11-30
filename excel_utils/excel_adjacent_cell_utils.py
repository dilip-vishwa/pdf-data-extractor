
from excel_utils.excel_color_utils import check_if_given_cell_is_text_or_green_or_empty
from excel_utils.excel_direction_utils import get_cell_far_in_given_direction_of_the_given_cell
from openpyxl.cell.cell import Cell

label_already_parsed = []

def find_adjacent_non_green_cell(cell, direction, distance=2, sheet=None, max_parsing_times=None) -> Cell|None:
    if isinstance(cell, str):
        cell = sheet[cell]

    if cell.value is None:
        if direction == 'left':
            distance = distance + 2

        for i in range(1, distance):
            cell = get_cell_far_in_given_direction_of_the_given_cell(cell, direction, i, sheet, max_parsing_times)
            if cell is not None:
                type_of_cell = check_if_given_cell_is_text_or_green_or_empty(cell, sheet)
                if type_of_cell == 'empty':
                    continue
                elif type_of_cell == 'green':
                    continue
                elif type_of_cell == 'text':
                    return cell
        
        return None
    else:
        return cell

def rules_for_finding_adjacent_label_cell(cell, direction, distance=2, sheet=None, max_parsing_times=None):
    if cell is None:
        return None
    
    if isinstance(cell, str):
        cell = sheet[cell]
    
    if direction == 'left':
        # iterate on all the left cells till 
        #   either no text is found
        #   or any text is found
        cell, type_of_cell = find_about_adjacent_cell(cell, 'left', 2, sheet, max_parsing_times)
        # if green cell is encountered, then return None i.e. this left cell should not be considered as label
        # if text cell is encountered, then return that cell i.e. this left cell should considered as label
        # if empty cell is encountered, then continu6e till the end of the row i.e. no label found.
        if type_of_cell == 'green':
            return None
        elif type_of_cell == 'text':
            return cell.value
        elif type_of_cell == 'empty':
            return None

    elif direction == 'top':
        # if top cell is green and already parsed, then return None
        if check_if_given_cell_is_text_or_green_or_empty(cell, sheet) == 'green' and cell.coordinate in label_already_parsed:
            return None
        if check_if_given_cell_is_text_or_green_or_empty(cell, sheet) == 'text':
            # if right cell is green and already parsed, then return None
            if check_if_adjacent_cell_is_green_or_empty_or_text_cell(cell, 'right', sheet, max_parsing_times) == 'green':
                return None
            return cell.value

        cell = get_cell_far_in_given_direction_of_the_given_cell(cell, 'right', 1, sheet, max_parsing_times)
        cell, type_of_cell = find_about_adjacent_cell(cell, 'right', 2, sheet, max_parsing_times)
        # using top cell, look for the right cell
        #   if green cell is encountered, then return None i.e. this top cell should not be considered as label
        #   if text cell is encountered, then return top Cell i.e. this top cell should considered as label
        #   if empty cell is encountered, then continue for 2 more cells. If found, then consider that cell as label
        if type_of_cell == 'green':
            return None
        elif type_of_cell == 'text':
            return cell.value
        elif type_of_cell == 'empty':
            return None

        # using top cell, look for the left cell
        #   No action required

        # using top cell, look for the bottom cell
        #   No action required

        # using top cell, look for the top cell
        #   No action required
    elif direction == 'right':
        # check if cell is having text, then do not consider this cell as label
        if check_if_given_cell_is_text_or_green_or_empty(cell, sheet) == 'text':
            return None
        cell, type_of_cell = find_about_adjacent_cell(cell, 'right', 2, sheet, max_parsing_times)
        # using right cell, look for the right cell (do 2-3 iterations for searching on right cell)
        #   if green cell is encountered, then return None i.e. this right cell should not be considered as label
        #   if text cell is encountered, then return right Cell i.e. this right cell should considered as label
        #   if empty cell is encountered, then continue for 2 more cells. If found, then consider that cell as label
        if type_of_cell == 'green':
            return None
        elif type_of_cell == 'text':
            # check again if it has any green cell right to the current cell
            if check_if_adjacent_cell_is_green_or_empty_or_text_cell(cell, 'right', sheet, max_parsing_times) == 'green':
                return None
            return cell.value
        elif type_of_cell == 'empty':
            return None
        # using right cell, look for the top cell
        #   No action required

        # using right cell, look for the bottom cell
        #   No action required

        # using right cell, look for the right cell
        #   No action required
    elif direction == 'bottom':
        cell, type_of_cell = find_about_adjacent_cell(cell, 'bottom', 2, sheet, max_parsing_times)
        # Nothing as of now
    else:
        pass


def find_about_adjacent_cell(cell, direction, distance=2, sheet=None, max_parsing_times=None) -> Cell|None:
    if isinstance(cell, str):
        cell = sheet[cell]

    if direction == 'left':
        distance = distance + 2

    for i in range(1, distance):
        if cell is not None:
            type_of_cell = check_if_given_cell_is_text_or_green_or_empty(cell, sheet)
            if type_of_cell == 'empty':
                cell = get_cell_far_in_given_direction_of_the_given_cell(cell, direction, i, sheet, max_parsing_times)
                continue
            elif type_of_cell == 'green':
                return cell, 'green'
            elif type_of_cell == 'text':
                return cell, 'text'
    
    return None, 'empty'



def find_adjacent_non_empty_cell(cell, direction, distance=2, sheet=None, max_parsing_times=None) -> Cell|None:
    if isinstance(cell, str):
        cell = sheet[cell]

    if cell.value is None:
        if direction == 'left':
            distance = distance + 2

        for i in range(1, distance):
            cell = get_cell_far_in_given_direction_of_the_given_cell(cell, direction, i, sheet, max_parsing_times)
            if cell is not None:
                type_of_cell = check_if_given_cell_is_text_or_green_or_empty(cell, sheet)
                if type_of_cell == 'empty':
                    continue
                elif type_of_cell == 'green':
                    return cell
                elif type_of_cell == 'text':
                    return cell
        
        return None
    else:
        return cell

def check_if_adjacent_cell_is_green_or_empty_or_text_cell(cell, direction, sheet=None, max_parsing_times=None):
    if isinstance(cell, str):
        cell = sheet[cell]

    cell = get_cell_far_in_given_direction_of_the_given_cell(cell, direction, 1, sheet, max_parsing_times)
    return check_if_given_cell_is_text_or_green_or_empty(cell, sheet)


def find_label_of_green_cell_using_rules(green_cell, sheet=None, max_parsing_times=None):
    left_cell, top_cell, right_cell, bottom_cell = get_cell_far_in_given_direction_of_the_given_cell(green_cell, 'all', 1, sheet, max_parsing_times)
    # print(green_cell.coordinate, "neighbouring cells", left_cell.value if left_cell else None, top_cell.value if top_cell else None, right_cell.value if right_cell else None, bottom_cell.value if bottom_cell else None)
    left_cell_label = rules_for_finding_adjacent_label_cell(left_cell, 'left', 2, sheet, max_parsing_times)
    top_cell_label = rules_for_finding_adjacent_label_cell(top_cell, 'top', 2, sheet, max_parsing_times)
    right_cell_label = rules_for_finding_adjacent_label_cell(right_cell, 'right', 2, sheet, max_parsing_times)
    bottom_cell_label = rules_for_finding_adjacent_label_cell(bottom_cell, 'bottom', 2, sheet, max_parsing_times)
    # if left label is None and top label is None, then take immediate top label as top label
    if left_cell_label is None:
        top_cell = green_cell
        
        if left_cell is None:
            left_cell = green_cell

        # get left cell value
        while left_cell_label is None:
            left_cell = get_cell_far_in_given_direction_of_the_given_cell(left_cell, 'left', 1, sheet, max_parsing_times)
            left_cell, type_of_cell = find_about_adjacent_cell(left_cell, 'left', 6, sheet, max_parsing_times)
            if type_of_cell == 'text':
                left_cell_label = left_cell.value
                break
            if left_cell is None:
                break
        # get top cell value
        while top_cell_label is None:
            top_cell = get_cell_far_in_given_direction_of_the_given_cell(top_cell, 'top', 1, sheet, max_parsing_times)
            top_cell, type_of_cell = find_about_adjacent_cell(top_cell, 'top', 6, sheet, max_parsing_times)
            if type_of_cell == 'text':
                top_cell_label = top_cell.value
                break
            if top_cell is None:
                break
    # print(green_cell.coordinate, "neighbouring cells after applying rules", left_cell_label, top_cell_label, right_cell_label, bottom_cell_label)
    label_already_parsed.append(green_cell.coordinate)
    return [green_cell.coordinate, left_cell_label, top_cell_label, right_cell_label, bottom_cell_label]


def find_labels_of_given_empty_green_cell_to_be_filled(green_cell, sheet=None, max_parsing_times=None):
    left_cell, top_cell, right_cell, bottom_cell = get_cell_far_in_given_direction_of_the_given_cell(green_cell, 'all', 1, sheet, max_parsing_times={'left': 2, 'top': 2, 'right': 2, 'bottom': 2})


    try: # For left cell
        type_of_cell = check_if_given_cell_is_text_or_green_or_empty(left_cell)
        if type_of_cell == 'empty':
            left_cell = find_adjacent_non_empty_cell(left_cell, 'left', 4, sheet=sheet, max_parsing_times=max_parsing_times)

        if type_of_cell == 'green':
            if check_if_adjacent_cell_is_green_or_empty_or_text_cell(left_cell, 'left', sheet, max_parsing_times) == 'text':
                pass
    except Exception as e:
        left_cell = None


    try: # For top cell
        type_of_cell = check_if_given_cell_is_text_or_green_or_empty(top_cell)
        if type_of_cell == 'empty':
            # if not (check_if_adjacent_cell_is_green_or_empty_or_text_cell(top_cell, 'right', sheet, max_parsing_times) == 'green' 
            #     or check_if_adjacent_cell_is_green_or_empty_or_text_cell(top_cell, 'right', sheet, max_parsing_times) == 'text'):
            #     top_cell = find_adjacent_non_empty_cell(top_cell, 'top', 3, sheet=sheet)
            # else:
            #     top_cell = None
            top_cell = top_cell
        elif type_of_cell == 'green':
            if check_if_adjacent_cell_is_green_or_empty_or_text_cell(top_cell, 'left', sheet, max_parsing_times) == 'text':
                pass
        elif type_of_cell == 'text':
            adjacent_cell_type = check_if_adjacent_cell_is_green_or_empty_or_text_cell(top_cell, 'right', sheet, max_parsing_times)
            if adjacent_cell_type == 'green':
                top_cell = None
            elif adjacent_cell_type == 'text':
                top_cell = top_cell
            elif adjacent_cell_type == 'empty':
                top_cell = top_cell
            
                
    except Exception as e:
        top_cell = None


    try: # For right cell
        type_of_cell = check_if_given_cell_is_text_or_green_or_empty(right_cell)
        if type_of_cell == 'empty':
            # if not (check_if_adjacent_cell_is_green_or_empty_or_text_cell(right_cell, 'right', sheet, max_parsing_times) == 'green' 
            #     or check_if_adjacent_cell_is_green_or_empty_or_text_cell(right_cell, 'right', sheet, max_parsing_times) == 'text'):
            #     right_cell = find_adjacent_non_green_cell(right_cell, 'right', 3, sheet=sheet)
            # else:
            #     right_cell = None

            adjacent_cell_type = check_if_adjacent_cell_is_green_or_empty_or_text_cell(right_cell, 'right', sheet, max_parsing_times)
            if adjacent_cell_type == 'green':
                right_cell = None
            elif adjacent_cell_type == 'text':
                right_cell = right_cell
            elif adjacent_cell_type == 'empty':
                right_cell = right_cell

        elif type_of_cell == 'green':
            if check_if_adjacent_cell_is_green_or_empty_or_text_cell(right_cell, 'left', sheet, max_parsing_times) == 'text':
                pass
    except Exception as e:
        right_cell = None


    try: # For bottom cell
        type_of_cell = check_if_given_cell_is_text_or_green_or_empty(bottom_cell)
        if type_of_cell == 'empty':
            if not (check_if_adjacent_cell_is_green_or_empty_or_text_cell(bottom_cell, 'bottom', sheet, max_parsing_times) == 'green' 
                or check_if_adjacent_cell_is_green_or_empty_or_text_cell(bottom_cell, 'bottom', sheet, max_parsing_times) == 'text'):
                bottom_cell = find_adjacent_non_green_cell(bottom_cell, 'bottom', 3, sheet=sheet)
            else:
                bottom_cell = None
        elif type_of_cell == 'green':
            bottom_cell = None
        elif type_of_cell == 'text':
            bottom_cell = None
    except Exception as e:
        bottom_cell = None



    if left_cell is None:
        left_cell = Cell(value=None, worksheet=sheet)
    if top_cell is None:
        top_cell = Cell(value=None, worksheet=sheet)
    if right_cell is None:
        right_cell = Cell(value=None, worksheet=sheet)
    if bottom_cell is None:
        bottom_cell = Cell(value=None, worksheet=sheet)
    
    print(green_cell.coordinate, left_cell.value, top_cell.value, right_cell.value, bottom_cell.value)
    label_already_parsed.append(green_cell.coordinate)
    return left_cell.value, top_cell.value, right_cell.value, bottom_cell.value
