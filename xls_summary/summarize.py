import os
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from xlrd import open_workbook

if __name__ == '__main__':
    # List all excel files in current folder
    my_dir = '.'
    template_name = 'summary_template.xlsx'
    summary_name = 'summary.xlsx'
    cur_filelist = [f for f in os.listdir(my_dir)
                    if os.path.isfile(os.path.join(my_dir, f))
                    and (f.endswith(".xls") or f.endswith(".xlsx"))]
    if template_name not in cur_filelist:
        print("Can't find a template, aborting")
        quit()
    cur_filelist.remove(template_name)
    # Create a summary file if it doesn't exist, otherwise load the current
    # one and append to it
    if summary_name in cur_filelist:
        # Load it up, then remove from cur_filelist
        wb_s = load_workbook(os.path.join(my_dir, summary_name))
        ws_s = wb_s.active
        cur_filelist.remove(summary_name)
    else:
        wb_s = Workbook()
        ws_s = wb_s.active
        ws_s.title = "Summary of data"
    # Scan template file to find indicators
    wb_t = load_workbook(os.path.join(my_dir, template_name))
    locations = {}
    for sheet_name in wb_t.get_sheet_names():
        sheet = wb_t.get_sheet_by_name(sheet_name)
        cells = sheet.get_cell_collection()
        for c in cells:
            if isinstance(c.value, str):
                if c.value.startswith('::') and c.value.endswith('::'):
                    row_identifier = c.value[2:-2]
                    locations[row_identifier] = {'sheet_name': sheet_name,
                                                 'col': c.column,
                                                 'row': c.row}
    for f in cur_filelist:
        # Open file
        wb = open_workbook(os.path.join(my_dir, f))
        summary_row = {}
        for summary_col, identifier in locations.items():
            # Extract values at locations indicated in template file
            sheet = wb.sheet_by_name(identifier['sheet_name'])
            col = column_index_from_string(identifier['col']) - 1
            row = identifier['row'] - 1
            val = sheet.cell_value(row, col)
            summary_row[summary_col] = val
        print(summary_row)
        # Create new row in summary file (check for old?)
    wb_s.save(filename=os.path.join(my_dir, summary_name))
