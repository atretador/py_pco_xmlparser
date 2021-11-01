# required modules
import xml.etree.cElementTree as ET
import glob, xlrd, os


# parse sheets
def sheet_to_xml(data):
    sheet_file = xlrd.open_workbook_xls(data)
    sheetn = sheet_file.sheets()
    root = ET.Element('root')
    processo = ET.SubElement(root, 'processo')
    processo.set('idProcesso', data[0:4])
    processo.set('dtMovimento', data[5:7] + "/" + data[7:9] + "/" + data[9:13])
    itens = ET.SubElement(processo, 'itens')

    # goes thru all the sheets on the file
    for sht in range(0, len(sheetn)):
        sheet = sheet_file.sheet_by_index(0)
        rows = sheet.nrows
        # goes thru all the rows in the file
        # for row in range(1, 2):
        for row in range(1, rows):
            print(row)
            item = ET.SubElement(itens, 'item')
            item.set('dtplanej', str(sheet.cell_value(row, 0)))
            item.set('valor', str(sheet.cell_value(row, 1)))
            item.set('CO', str(int(sheet.cell_value(row, 2))))
            item.set('CLASSE', str(sheet.cell_value(row, 3)))
            item.set('OPER', str(sheet.cell_value(row, 4)))
            item.set('CC', str(int(sheet.cell_value(row, 5))))
            item.set('ITCTB', str(sheet.cell_value(row, 6)))
            item.set('CLVLR', str(sheet.cell_value(row, 7)))
            item.set('IDREF', str(int(sheet.cell_value(row, 8))))
            item.set('TIPO', str(int(sheet.cell_value(row, 9))))
    xml_file = ET.tostring(root)
    name = data[:-4] + ".xml"
    ret = [name, xml_file]
    return ret


def xml_to_file(ret_obj):
    f = open(path+"//export//"+ret_obj[0], "wb")
    text = ret_obj[1]
    f.write(text)
    f.close()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print("start")
    path = os.getcwd()
    os.chdir(path+"/import")
    print("files")
    # get all spreadsheets on directory
    for file in glob.glob("*.xls"):
        xml_to_file(sheet_to_xml(file))