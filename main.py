import sys

if __name__ == '__main__':
    option_list = ["", "open", "simple", "", "", ""]
    if sys.argv.__len__() == 1:
        l = option_list.__len__()
        sys.stdout.write(str(option_list[1:l]))
    elif sys.argv[1] == option_list[1]:
        from almighty_xlrd.workbook_manipulation import open_workbook_my

        open_workbook_my()
    elif sys.argv[1] == option_list[2]:
        from almighty_xlrd.workbook_manipulation import simple as workbook_simple

        workbook_simple()
    elif sys.argv[1] == option_list[3]:
        pass
    elif sys.argv[1] == option_list[4]:
        pass
    elif sys.argv[1] == option_list[5]:
        pass
    else:
        sys.stdout.write(option_list[1:])
