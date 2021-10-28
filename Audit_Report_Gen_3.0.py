from openpyxl import load_workbook
from openpyxl import Workbook
import re

def matching(wb1, wb2, wb_out, tail, order1, order2, output_order):
    # wb_out = Workbook()
    for j in range(len(order1)):
        sheet1 = wb1[order1[j]]
        sheet2 = wb2[order2[j]]
        ws = wb_out.create_sheet(output_order[j])
        i = 0
        for row1 in sheet1.iter_rows(max_col=tail, values_only=True):
            cell_output = cell_data(row1)
            match = False
            if cell_output[0] == 'none':
                break
            for row2 in sheet2.iter_rows(max_col=2, values_only=True):
                cell_temp = cell_data(row2)
                if str(cell_output[0]).lower() == str(cell_temp[0]).lower():
                    match = True
                    cell_output.append(cell_temp[1])
                    ws.append(cell_output)
                    break
            if match == False:
                cell_output.append('none')
                ws.append(cell_output)
            i = i + 1
        print(output_order[j] + ': ' + str(i-1))

def cell_data(row):
        cells = []
        for cell in row:
            if cell is None:
                cells.append('none')
            else:
                cells.append(cell)
        return cells

#--------------------------------------------- Input Settings -------------------------------------------------
tail = 2
Week = 43
comp = Week - 1
order1 = ['Main', 'STR', 'SORP']    # This week
order2 = ['Main', 'STR', 'SORP']    # Last week
output_order = ['Main', 'STR', 'SORP']
# order2 = ['Main', 'STR', 'Main']    # w-2 week
# output_order = ['Main', 'STR', 'Production']

folder = 'C:\\Users\\Yvonne\\Documents\\Results\\'
#--------------------------------------------------------------------------------------------------------------

wb1 = load_workbook(folder + '\\All\\2021_W' +  str(Week)  + '.xlsx')
wb2 = load_workbook(folder + '\\All\\2021_W' + str(comp) + '.xlsx')

wb_out = Workbook()
matching(wb1, wb2, wb_out, tail, order1, order2, output_order)
wb_out.save(folder + '\\Audit_Report\\Audit_Report_W' + str(Week) + '.xlsx')

# wb_out.save(folder + '\\Audit_Report\\Audit_Report_Ww' + str(Week) + '.xlsx')