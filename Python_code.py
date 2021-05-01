import openpyxl
#purpose of function is to allow the delimiter to function even if the delimiter is not available at the position specified; also handles when row is empty
def split_delimiter(filename, ini_col, ini_row, delimiter, number_from_left, output_name): #number from left; first number is 0
    wb = openpyxl.load_workbook(filename)
    sheet_interest = wb.sheetnames[0]  # select first sheet
    sheet = wb[sheet_interest]
    max_row = sheet.max_row
    contents_of_cells = []
    left_content = []
    right_content = []
    delimiter_position =[]
    for i in range(0, max_row - ini_row+1):
        if sheet.cell(row=ini_row + i, column=ini_col).value == None:
            contents_of_cells = contents_of_cells + ['']
        else:
            contents_of_cells = contents_of_cells + [sheet.cell(row=ini_row + i, column=ini_col).value]
    for j in range(0, max_row - ini_row+1):
        len_match = len([pos for pos, char in enumerate(contents_of_cells[j]) if char == delimiter])
        if len_match <= number_from_left:
            delimiter_position = delimiter_position + [-1]
            print('row ', j + ini_row, ' in the original Excel file has no delimiter at position of interest')
        else:
            delimiter_position = delimiter_position + [[pos for pos, char in enumerate(contents_of_cells[j]) if char == delimiter][number_from_left]]
    for k in range(0, max_row - ini_row+1):
        len_string = len(contents_of_cells[k])
        delimiter_value = delimiter_position[k]
        if delimiter_value == -1:
            left_content = left_content + [contents_of_cells[k]]
            right_content = right_content + ['']
        elif contents_of_cells[k][delimiter_value+1] == ' ':
            left_content = left_content + [contents_of_cells[k][0:delimiter_value]]
            right_content = right_content + [contents_of_cells[k][delimiter_value + 2:len_string]]
        else:
            left_content = left_content + [contents_of_cells[k][0:delimiter_value]]
            right_content = right_content + [contents_of_cells[k][delimiter_value+1:len_string]]
    wb_new = openpyxl.Workbook()
    sheet = wb_new['Sheet']
    sheet.cell(row=1, column=1).value = 'Left of delimiter'
    sheet.cell(row=1, column=2).value = 'Right of delimiter'
    for x in range(1, max_row - ini_row+2):
        sheet.cell(row=x + 1, column=1).value = left_content[x-1]
        sheet.cell(row=x + 1, column=2).value = right_content[x-1]
    wb_new.save(output_name)
