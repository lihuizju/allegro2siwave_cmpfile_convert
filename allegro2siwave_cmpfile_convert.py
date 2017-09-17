# -*- coding: utf-8 -*-
#  import sys
import xlrd
import pdb

_DEBUG = False


def open_excel(file='file.xlsx'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print(str(e))


def excel_table_byindex(file='file.xlsx', colnameindex=0, by_index=0):
    if _DEBUG is True:
        pdb.set_trace()
    data = open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows
    ncols = table.ncols
    colnames = table.row_values(colnameindex)
    part_ref_index = 0
    MPN_index = 0
    for colnum in range(0, ncols):
        if colnames[colnum] == "Part Reference":
            part_ref_index = colnum
        elif colnames[colnum] == "Manufacturer Part Number":
            MPN_index = colnum
    list = []
    #  app = {}
    for rownum in range(colnameindex + 1, nrows):
        row = table.row_values(rownum)
        if row:
            app = {}
            app[colnames[part_ref_index]] = row[part_ref_index]
            if row[MPN_index] == '':
                row[MPN_index] = 'NM'
            app[colnames[MPN_index]] = row[MPN_index]
            list.append(app)
    return list


def cmp_file_read(file='cmpfile.cmp'):
    try:
        cmp_file = open(file, 'r')
        _list_content = []
        line = cmp_file.readline()
        while line:
            _list_content.append(line)
            line = cmp_file.readline()
        return _list_content
    finally:
        if cmp_file:
            cmp_file.close()


def replace_part_no(_content, pn_table):
    str_begin = ['B_CAP', 'B_RES', 'B_IND', 'B_IC']
#    str_end = ['E_CAP', 'E_RES', 'E_IND', 'E_IC']
    replaced_content = _content
    wrong_pn = ''
    right_pn = ''
    for i in range(0, len(_content)):
        if any(x in _content[i] for x in str_begin):
            line_split = _content[i].split()
            part_ref = line_split[1].strip('"')
            wrong_pn = line_split[2]
            for j in range(0, len(pn_table)):
                if pn_table[j]['Part Reference'] == part_ref:
                    # pdb.set_trace()
                    right_pn = '"' + pn_table[j]['Manufacturer Part Number'] + '"'
                    replaced_content[i] = _content[i].replace(wrong_pn, right_pn)
                    break
                    # del pn_table[j]
        wrong_pn = wrong_pn.strip('"')
        right_pn = right_pn.strip('"')
        replaced_content[i] = _content[i].replace(wrong_pn, right_pn)
    return replaced_content


def main():
    tables = excel_table_byindex("E:\Workspace\Cadence\Cxx\\bom.xlsx")
#    if _DEBUG is True:
#        import pdb
#        pdb.set_trace()
    #  for row in tables:
    #    print(row)
    content = cmp_file_read("E:\Workspace\Cadence\Cxx\Cxx.cmp")
    _content = replace_part_no(content, tables)
    with open("E:\Workspace\Cadence\Cxx\Cxx_new.cmp", 'w') as _file:
        _file.writelines(_content)
    #print(content)


if __name__ == "__main__":
    main()
