# create a dictionary of rules like {"ch":'щ'}
# presume data stored in a excel-'file' with known 'sheetname'


def get_rules(file="./rules.xlsx", sheetname='rules'):
    result = {}
    # open excel-'file'
    # use excel library 'openpyxl'
    import openpyxl

    wb = openpyxl.load_workbook(file)
    print('List of sheets: ', wb.get_sheet_names())

    wb.close()
    return result


if __name__ == '__main__':
    print("Start...")
    rules = get_rules()
    print(rules)
    if(rules):
        print('Rules are not empty.')
    print('End')
