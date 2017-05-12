import os
from openpyxl import Workbook
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
    print(cur_filelist)
    # Create a summary file if it doesn't exist, otherwise load the current
    # one and append to it
    if summary_name in cur_filelist:
        # Load it up, then remove from cur_filelist
        pass
    else:
        pass
    # Scan template file to find indicators
    for f in cur_filelist:
        # Open file
        # Extract values at locations indicated in template file
        # Create new row in summary file (check for old?)
        pass
