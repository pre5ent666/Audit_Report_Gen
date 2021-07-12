from openpyxl import load_workbook
from openpyxl import Workbook
import re

def matching(wb1, wb2, Week, tail, order1, order2, output_order):
    wb_out = Workbook()
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
    wb_out.save(folder + 'Audit_Report_W' + str(Week) + '.xlsx')

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
Week = 28
order1 = ['Main', 'Product', 'STR', 'Product', 'Main']    # This week
order2 = ['Main', 'Product', 'STR', 'Main', 'Product']    # Last week
output_order = ['Main', 'Product', 'STR', 'W28P_W27M', 'W28M_W27P']

folder = 'C:\\Users\\Yvonne\\Documents\\Results\\'
#--------------------------------------------------------------------------------------------------------------

wb1 = load_workbook(folder + '2021_W' +  str(Week)  + '.xlsx')
wb2 = load_workbook(folder + '2021_W' + str(Week-1) + '.xlsx')

matching(wb1, wb2, Week, tail, order1, order2, output_order)