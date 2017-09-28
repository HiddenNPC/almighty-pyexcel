def open_workbook_my():
    """
    Opening Workbooks
    Workbooks can be loaded either from a file, an mmap.mmap object or from a string
    :return:none
    """

    from mmap import mmap, ACCESS_READ
    from xlrd import open_workbook
    print open_workbook('simple.xls')
    with open('simple.xls', 'rb') as f:
        print open_workbook(
            file_contents=mmap(f.fileno(), 0, access=ACCESS_READ)
        )
    aString = open('simple.xls', 'rb').read()
    print aString
    print open_workbook(file_contents=aString)


def simple():
    '''
    Navigating a Workbook
    Here is a simple example of workbook navigation
    :return:none
    '''
    from xlrd import open_workbook
    wb = open_workbook('simple.xls')
    for s in wb.sheets():
        print 'Sheet:', s.name
        for row in range(s.nrows):
            values = []
            for col in range(s.ncols):
                values.append(s.cell(row, col).value)
            print values
        print
