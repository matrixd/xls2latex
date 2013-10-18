import xlrd

xls = 'test/1.xls'

b = xlrd.open_workbook(xls)


def create_table(rows, **kwargs):
    table = "\\begin{tabular}{ |"
    width = kwargs.get('right')-kwargs.get('left')+1
    height = kwargs.get('bottom')-kwargs.get('top')
    for i in range(width):
        table += ' c | '
    table += '}\n   \\hline '
    for row in rows:
        #proccessing on table rows
        row = row[0]
        ncell = kwargs.get('left')
        for c in row:
            if c[1] == ncell:
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
    #closing table
    table = table[:-10]
    table += '\\hline \\end{tabular}'
    print(table)


def read_table(book):
    rows = []
    right = 0
    left = 0
    top = 0
    bottom = 0
    first = True

    for nrow in range(book.sheet_by_index(0).nrows):
        row = book.sheet_by_index(0).row(nrow)
        print(row)
        print('row %d' % nrow)
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

    create_table(rows, left=left, right=right, top=top, bottom=bottom)

read_table(b)