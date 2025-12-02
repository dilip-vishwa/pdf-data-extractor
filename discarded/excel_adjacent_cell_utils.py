
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
