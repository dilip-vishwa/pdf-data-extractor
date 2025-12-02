

def get_border_present_or_not_around_cell(cell, sheet=None):
    """
    Check if a cell has any border present around it. Output is in tulple format like (left, right, top, bottom) in boolean format
    """
    if isinstance(cell, str):
        cell = sheet[cell]

    border = cell.border
    if border:
        return (border.left and border.left.style, 
           border.right and border.right.style,
           border.top and border.top.style,
           border.bottom and border.bottom.style)

    return (None, None, None, None)
