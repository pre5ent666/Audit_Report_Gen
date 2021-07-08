from openpyxl import load_workbook
from openpyxl import Workbook
import re

def matching(wb1, wb2, Week, tail, order1, order2, output_order):
    # order1 = ['Main', 'Product', 'Product', 'Main']
    # order2 = ['Main', 'Product', 'Point', 'STR']
    # output_order = ['Main', 'Product', 'W25P_Point', 'W25M_STR']
    wb_out = Workbook()
    for j in range(len(order1)):
        sheet1 = wb1[order1[j]]
        sheet2 = wb2[order2[j]]
        ws = wb_out.create_sheet(output_order[j])
        i = 1
        for row1 in sheet1.iter_rows(max_col=tail, values_only=True):
            cell_output = cell_data(row1)
            match = False
            for row2 in sheet2.iter_rows(max_col=2, values_only=True):
                cell_temp = cell_data(row2)
                if str(cell_output[0]).lower() == str(cell_temp[0]).lower():
                    match = True
                    cell_output.append(cell_temp[1])
                    ws.append(cell_output)
                    break
            if match == False:
                print(cell_output)
                ws.append(cell_output)
            print(i)
            i = i + 1
    wb_out.save(folder + 'Audit_Report_W' + str(Week) + '.xlsx')

def cell_data(row):
        cells = []
        for cell in row:
            if cell is None:
                cells.append('none')
            else:
                cells.append(cell)
        return cells

tail = 2
Week = 27
order1 = ['Main', 'Product', 'STR']
order2 = ['Main', 'Product', 'STR']
output_order = ['Main', 'Product', 'STR']
folder = 'C:\\Users\\Yvonne\\Documents\\Results\\'
# folder = 'C:\\Users\\Yvonne\\Downloads\\'

wb1 = load_workbook(folder + '2021_W' + str(Week-1) + '.xlsx')
wb2 = load_workbook(folder + '2021_W' +  str(Week)  + '.xlsx')

matching(wb1, wb2, Week, tail, order1, order2, output_order)