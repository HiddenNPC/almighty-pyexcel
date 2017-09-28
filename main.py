import sys

if __name__ == '__main__':
    option_list = ["", "open", "simple", "introspect_book", "introspect_sheet", "xlsxCreate"]
    l = option_list.__len__()
    if sys.argv.__len__() == 1:
        sys.stdout.write(str(option_list[1:l]))
        from almighty_xlrd.xls_manipulation import saveExcel

        saveExcel()
    elif sys.argv[1] == option_list[1]:
        from almighty_xlrd.xls_manipulation import open_xls

        open_xls()
    elif sys.argv[1] == option_list[2]:
        from almighty_xlrd.xls_manipulation import simple as workbook_simple

        workbook_simple()
    elif sys.argv[1] == option_list[3]:
        from almighty_xlrd.xls_manipulation import introspect_book

        introspect_book()
    elif sys.argv[1] == option_list[4]:
        from almighty_xlrd.xls_manipulation import introspect_sheet

        introspect_sheet()
    elif sys.argv[1] == option_list[5]:
        from almighty_xlrd.xls_manipulation import xlsxCreate

        xlsxCreate()
    else:
        sys.stdout.write(str(option_list[1:l]))
