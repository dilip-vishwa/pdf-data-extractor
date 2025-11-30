import openpyxl

def get_excel_sheet(file_path):
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        # Assuming we want to check all sheets, or just the active one? 
        # User didn't specify, but usually active sheet is the main one.
        sheet = wb.active
        print(f"Processing sheet: {sheet.title}")
        
        return wb, sheet
    except Exception as e:
        print(f"Error: {e}")
        return None
