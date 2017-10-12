def open_xls():
    """
    Opening Workbooks
    Workbooks can be loaded either from a file, an mmap.mmap object or from a string
    :return:none
    """

    from mmap import mmap, ACCESS_READ
    from xlrd import open_workbook
    print open_workbook('simple.xls')
    with open('simple.xlsx', 'rb') as f:
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
    wb = open_workbook('simple.xlsx')
    for s in wb.sheets():
        print 'Sheet:', s.name
        for row in range(s.nrows):
            values = []
            for col in range(s.ncols):
                values.append(s.cell(row, col).value)
            print values
        print


def introspect_book():
    from xlrd import open_workbook
    book = open_workbook('simple.xls')
    print book.nsheets
    for sheet_index in range(book.nsheets):
        print book.sheet_by_index(sheet_index)
    print book.sheet_names()
    for sheet_name in book.sheet_names():
        print book.sheet_by_name(sheet_name)
    for sheet in book.sheets():
        print sheet


def introspect_sheet():
    from xlrd import open_workbook, cellname
    book = open_workbook('odd.xls')
    sheet = book.sheet_by_index(0)
    print sheet.name
    print sheet.nrows
    print sheet.ncols
    for row_index in range(sheet.nrows):
        for col_index in range(sheet.ncols):
            print cellname(row_index, col_index), '-',
    print sheet.cell(row_index, col_index).value


def saveExcel():
    from tempfile import TemporaryFile
    from xlwt import Workbook
    book = Workbook()
    sheet1 = book.add_sheet('Sheet 1')
    book.add_sheet('Sheet 2')
    sheet1.write(0, 0, 'A1')
    sheet1.write(0, 1, 'B1')
    row1 = sheet1.row(1)
    row1.write(0, 'A2')
    row1.write(1, 'B2')
    sheet1.col(0).width = 10000
    sheet2 = book.get_sheet(1)
    sheet2.row(0).write(0, 'Sheet 2 A1')
    sheet2.row(0).write(1, 'Sheet 2 B1')
    sheet2.flush_row_data()
    sheet2.write(1, 0, 'Sheet 2 A3')
    sheet2.col(0).width = 5000
    sheet2.col(0).hidden = True
    book.save('simple.xls')
    book.save(TemporaryFile())


def demo_xlsxCreate():
    ##############################################################################
    #
    # A simple example of some of the features of the XlsxWriter Python module.
    #
    # Copyright 2013-2017, John McNamara, jmcnamara@cpan.org
    #
    import xlsxwriter

    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook('demo.xlsx')
    worksheet = workbook.add_worksheet()

    # Widen the first column to make the text clearer.
    worksheet.set_column('A:A', 20)

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})

    # Write some simple text.
    worksheet.write('A1', 'Hello')

    # Text with formatting.
    worksheet.write('A2', 'World', bold)

    # Write some numbers, with row/column notation.
    worksheet.write(2, 0, 123)
    worksheet.write(3, 0, 123.456)

    # Insert an image.
    worksheet.insert_image('B5', 'logo.png')

    workbook.close()


def generate_xlsx_from_mongo(overview_db, table_list):
    import xlsxwriter
    result = []
    for table_name in table_list:
        cur_doc = overview_db['all_table_header_info'].find_one({"table_name": table_name})
        # Create an new Excel file and add a worksheet.
        workbook = xlsxwriter.Workbook('./private/%s.xlsx' % table_name)
        worksheet = workbook.add_worksheet()

        # Add a bold format to use to highlight cells.
        format_bold = workbook.add_format({'bold': True})

        # Widen the first column to make the text clearer.
        worksheet.set_column('A:GG', 20, format_bold)
        if cur_doc == None or len(cur_doc) == 0:
            return result
        else:
            result = cur_doc.get('table_header')
            for i in result:
                worksheet.write(0, result.index(i), i)
            workbook.close()
