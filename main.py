import xlrd

def create_table(kwargs):
    """
    method for creating latex code
    it returns string
    kwargs are = rows, left, right, top, bottom
    """
    table = "\\begin{tabular}{ |"
    width = kwargs.get('right')-kwargs.get('left')+1
    height = kwargs.get('bottom')-kwargs.get('top')
    rows = kwargs.get('rows')
    merged = kwargs.get('merged')
    for i in range(width):
        table += ' c | '
    table += '}\n   \\hline '
    currow = kwargs.get('top')
    for row in rows:
        #proccessing on table rows
        row = row[0]
        ncell = kwargs.get('left')
        for c in row:
            if c[1] == ncell:
                if merged and (currow, ncell) in merged:
                    l = merged[(currow, ncell)][3] - merged[(currow, ncell)][2]
                    table += '\\multicolumn{' + str(l) + '}{|c|}{' + str(c[0].value) + '}'
                    ncell += l-2
                else:
                    table += str(c[0].value) + ' & '

            else:
                #writing empty cols if necessery
                for k in range(c[1] - ncell):
                    table += ' & '
                    ncell += 1
            ncell += 1
        if row[-1][1] < kwargs.get('right'):
            for k in range(kwargs.get('right') - row[-1][1]):
                    table += ' & '
        table = table[:-2]
        table += "\\" + "\\" + "\n    \\hline "  #row end
        currow += 1

    #closing table
    table = table[:-10]
    table += '\\hline \\end{tabular}'
    return table


def read_table(book):
    """function for reading table from xlrd.book object"""
    rows = []
    right = 0
    left = 0
    top = 0
    bottom = 0
    first = True

    for nrow in range(book.sheet_by_index(0).nrows):
        row = book.sheet_by_index(0).row(nrow)
        cell_list = []
        ncell = 0
        for cell in row:
            if cell.ctype != 0:
                cell_list.append((cell, ncell))
                if first:
                    top = nrow
                    left = ncell
                    first = False
                if top > nrow:
                     top = nrow
                if ncell < left:
                    left = ncell
                if ncell > right:
                    right = ncell
                if nrow > bottom:
                    bottom = nrow
            ncell += 1
        if cell_list:
            rows.append((cell_list, nrow))

    #getting merged cells
    merged = {}
    for cell in book.sheet_by_index(0).merged_cells:
        merged.update({(cell[0], cell[2]): cell})

    return {'rows': rows, 'left': left, 'right': right, 'top': top, 'bottom': bottom, 'merged': merged}
    #return(rows, left, right, top, bottom)


if __name__ == "__main__":
    xls = input("Enter filename/path_to_file to convert")
    b = xlrd.open_workbook(xls, formatting_info=True)
    create_table(read_table(b))

