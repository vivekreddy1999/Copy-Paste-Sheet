from openpyxl import load_workbook

def copy_sheet_between_workbooks(source_path, target_path, sheet_name):
    
    # Load both workbooks
    source_wb = load_workbook(source_path)
    target_wb = load_workbook(target_path)
    
    # Get the sheet to copy from source workbook
    source_sheet = source_wb[sheet_name]
    
    # Create a new sheet in target workbook with the same name
    # If sheet already exists, we'll create one with a different name
    if sheet_name in target_wb.sheetnames:
        new_sheet_name = f"{sheet_name}_copy"
        target_sheet = target_wb.create_sheet(new_sheet_name)
    else:
        target_sheet = target_wb.create_sheet(sheet_name)
    
    # Copy cell values and styles
    for row in source_sheet.rows:
        for cell in row:
            # Get the cell coordinate (e.g., 'A1')
            target_cell = target_sheet[cell.coordinate]
            
            # Copy value
            target_cell.value = cell.value
            
            # Copy style
            if cell.has_style:
                target_cell.font = cell.font.copy()
                target_cell.border = cell.border.copy()
                target_cell.fill = cell.fill.copy()
                target_cell.number_format = cell.number_format
                target_cell.protection = cell.protection.copy()
                target_cell.alignment = cell.alignment.copy()
    
    # Copy column dimensions
    for col_letter, col_dim in source_sheet.column_dimensions.items():
        target_sheet.column_dimensions[col_letter].width = col_dim.width
    
    # Copy row dimensions
    for row_num, row_dim in source_sheet.row_dimensions.items():
        target_sheet.row_dimensions[row_num].height = row_dim.height
    
    # Copy merged cells
    for merged_cell_range in source_sheet.merged_cells.ranges:
        target_sheet.merge_cells(str(merged_cell_range))
    
    # Save the target workbook
    target_wb.save(target_path)
    
    # Close both workbooks
    source_wb.close()
    target_wb.close()

# usage
source_file = "Book1.xlsx"
target_file = "Book2.xlsx"
sheet_to_copy = "Sheet2"

copy_sheet_between_workbooks(source_file, target_file, sheet_to_copy)
