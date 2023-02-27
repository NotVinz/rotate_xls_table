# -*- coding: utf-8 -*-
import sys
import xlrd3
import xlwt

# Finds out if a cell is part of a group of merged cells
def isinmerged(row, column, merged_cells):
    for square in merged_cells:
        if square[0] <= row < square[1]:
            if square[2] <= column < square[3]:
                return True
    return False

def rotateworkbook(filename):
    inp_Workbook = xlrd3.open_workbook(filename, formatting_info=True)
    out_Workbook = xlwt.Workbook(encoding='utf-8')

    for nsheet in list(range(len(inp_Workbook.sheets()))):
        curr_inp_sheet = inp_Workbook.sheets()[nsheet]
        curr_out_sheet = out_Workbook.add_sheet(
            inp_Workbook.sheets()[nsheet].name)

        #style = xlwt.easyxf('align: rotation 90')
        style = xlwt.easyxf('align: rotation 90, vert centre, horiz centre')

        inp_sheet_nrows = curr_inp_sheet.nrows
        inp_sheet_ncols = curr_inp_sheet.ncols
        merged_cells = curr_inp_sheet.merged_cells
        orig_table = []
        for m in range(0, inp_sheet_nrows):
            cols = []
            for i in range(0, inp_sheet_ncols):
                cols.append(curr_inp_sheet.cell(m, i).value)
            orig_table.append(cols)

        rotate_table = []
        for t in range(0, inp_sheet_nrows):
            for s in range(0, inp_sheet_ncols):
                if not isinmerged(t, s, merged_cells):
                    curr_out_sheet.write(inp_sheet_ncols - s - 1, t, orig_table[t][s], style=style)
                else:
                    if False:
                        print("Skipped: " + str(t) + "," + str(s))
                        print("Val:")
                        print(orig_table[t][s])
                        print("")
                    pass

        for mgd in merged_cells:
            if False:
                print("Found group of merged cells:")
                print([inp_sheet_ncols - mgd[3], inp_sheet_ncols - mgd[2] - 1, mgd[0], mgd[1] - 1])

            curr_out_sheet.write_merge(inp_sheet_ncols - mgd[3], inp_sheet_ncols - mgd[2] - 1, mgd[0],
                                       mgd[1] - 1,
                                       orig_table[mgd[0]][mgd[2]],
                                       style=style)

    out_Workbook.save("rot_" + filename)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Something is wrong")
        print("Please run as Rotate_Excel.py <Filename.xls>")
        exit()
    # rotate('caca.xls')  # input file name
    filename = sys.argv[1]

    rotateworkbook(filename)
