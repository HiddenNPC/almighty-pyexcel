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
        from almighty_xlrd.xls_manipulation import generate_xlsx_from_mongo
        import configparser
        from tools.db_tools import get_db_client

        table_list = []
        my_config = configparser.ConfigParser()
        my_config.read('my_config.ini')
        overview_db = get_db_client(my_config)
        result = overview_db["all_table_header_info"].find({}, {"table_name": 1, "_id": 0})
        for table in result:
            table_list.append(table["table_name"].encode('ascii', 'ignore'))
        generate_xlsx_from_mongo(overview_db, table_list)
    else:
        sys.stdout.write(str(option_list[1:l]))
