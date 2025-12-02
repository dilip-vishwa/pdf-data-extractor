# PDF Data Extractor And Fill the Excel



## Following logic used to find the label of the cell to be filled

- iterate over all green found cells:

    - 1 - look for border around
        - if all corner present or not present, then no issue, we need to fill the data.
        - IMP: if border not present, then it is more sure that the label is for that particular cell.
        
    - 2 - if there is white cell while looking for label, then stop over there and go to next cell if not found at left, top, right and bottom.
    
    - 3 - Now try to find the label which might be present in left and maybe additonal in top, right or bottom.
        - if (left border is absent or present and) left label is there, then consider that as label.
            - if right label is also present, then look beyond that label if green cell is present, if yes, then leave that cell
                - if no cell after right label, then consider that also as a secondary label.
        - if top text is present, (then look if that is having same width else leave it. if same width,) then take that as secondary label only if that is not a label to another green cell.
            - if there is green cell on top, then go beyond that cell and repeat the same till the label is not found on left, top, right or at bottom.
        
        - if bottom is having green cell again, then look beyond that to get the label. If label is found, then consider it as label. 
            
        - IMP: if bottom text is present, then look if that is having same width else leave it. if same width, then take that as secondary or tertiary label
        IMP: if bottom cell is present, then look if that is having same width else leave it. if same width, then take that parse ahaead with that cell.

    
    - Note: Left label need to be parsed beyond 1-4 cell on the left, and it will surely get the left label else no label.
    
    - IMP: the parsing will always be from left to right all cells and then go to next row.

    - IMP: If cell is present next to another cell, then without top label or same width, then consider it as haing same left most label.(G54)
    
    - IMP: if right cell as no border on right with text on right, then consider it as label.
    
    - IMP IMP IMP: all logic is working except Impedance. Will look for it afterward.
    - IMP: for left, top, right and bottom label, we need to accept(consider) label occupying multiple rows or columns.