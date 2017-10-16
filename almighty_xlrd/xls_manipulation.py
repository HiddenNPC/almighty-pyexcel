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


def clearStrCharset(string):
    import re
    result_str = ""
    try:
        if string is None:
            result_str = ""
            return result_str
        elif not isinstance(string, basestring):
            return string
        else:
            result_str = str(string)
        result_str = re.sub(r'[^\x20-\x7e]', '', result_str)
        result_str = result_str.replace("\r", '').replace("\n", '').strip()
    except:
        # print [string]
        pass
    return result_str


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


def generateExcel_from_mongo(header, content, sheetName, file):
    import xlsxwriter
    print "start to save %s" % sheetName
    workbook = xlsxwriter.Workbook(file)
    worksheet = workbook.add_worksheet(sheetName)
    worksheet.set_column('A:GG', 20)
    formatHeader = workbook.add_format()
    formatYellow = workbook.add_format()
    formatRed = workbook.add_format()
    formatGreen = workbook.add_format()
    formatRow = workbook.add_format()
    formatYellow.set_bg_color("yellow")
    formatRed.set_bg_color("red")
    formatGreen.set_bg_color("green")
    formatHeader.set_bold()
    formatHeader.set_font_color('blue')
    formatHeader.set_align("center")
    formatHeader.set_align("vcenter")
    formatRow.set_bg_color("#daeef3")

    worksheet.write_row(0, 0, header, formatHeader)
    worksheet.freeze_panes(1, 0)
    cols = []
    try:
        for i in range(len(header)):
            cols.append(len(header[i]))
        for line in content:
            for i, values in enumerate(line):
                if values is not None:
                    if len(str(values)) > cols[i]:
                        cols[i] = len(str(values))

        for i in range(len(header)):
            if cols[i] > 80:
                cols[i] = len(header[i]) + 10
            worksheet.set_column(i, i, cols[i] + 1)
    except:
        # print "The column number is "
        pass
    # print content
    # worksheet.write_url(0,2,"internal:'Sheet1'!A1")
    hostname, number = "", 0
    for row_index in range(len(content)):
        if hostname != content[row_index][0]:
            number += 1
            hostname = content[row_index][0]
        if number % 2 == 0:
            worksheet.set_row(row_index + 1, None, formatRow)
            # for col_index in range(len(content[row_index])):
            #     worksheet.write(row_index+1,col_index,"",formatRow)

    for row_index in range(len(content)):
        for col_index in range(len(content[row_index])):
            try:
                content[row_index][col_index] = clearStrCharset(content[row_index][col_index])
                if isinstance(content[row_index][col_index], tuple):
                    if content[row_index][col_index][1] == "yellow":
                        worksheet.write(row_index + 1, col_index, str(content[row_index][col_index][0]), formatYellow)
                    elif content[row_index][col_index][1] == "red":
                        worksheet.write(row_index + 1, col_index, str(content[row_index][col_index][0]), formatRed)
                    elif content[row_index][col_index][1] == "green":
                        worksheet.write(row_index + 1, col_index, str(content[row_index][col_index][0]), formatGreen)
                else:
                    worksheet.write(row_index + 1, col_index, str(content[row_index][col_index]))
                    # worksheet.set_column(row_index+1,col_index,cols[row_index][col_index]+20)
            except:
                print "error", content[row_index], col_index
                worksheet.write(row_index + 1, col_index, "")
    print file
    workbook.close()
