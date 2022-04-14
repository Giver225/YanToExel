import sys
import yadisk
import xlwt
import getpass

dictionary_path = input('Введите путь к папке: ')
y = yadisk.YaDisk(token='AQAAAAARvKUuAAfHL8oVHsnSYEDXqjUh-t-mJcY')

def get_names_and_links(dictionary_path):
    global y
    try:
        dictionary_link = y.get_meta(dictionary_path)['public_url']
    except:
        y = yadisk.YaDisk(token='AQAAAAA6MxsrAAfHUiObZGAskkZBo76Id3MtFLM')
        dictionary_link = y.get_meta(dictionary_path)['public_url']
    names = []
    links = []

    gen = y.public_listdir(dictionary_link)
    for i in gen:
        y.publish(dictionary_path + i['path'])
    gen = y.public_listdir(dictionary_link)
    for i in gen:
        if y.get_public_type(i['public_url']) == 'dir':
            names_and_links = get_names_and_links(dictionary_path+i['path'])
            names += names_and_links[0]
            links += names_and_links[1]
        elif y.get_public_type(i['public_url']) == 'file':
            names.append(i['name'])
            links.append(i['public_url'])
    return names, links




result = get_names_and_links(dictionary_path)
print(result[0]) #names
print(result[1]) #links


book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet1")
for i in range(len(result[0])):
    sheet1.write(i, 0, result[0][i])
    sheet1.write(i, 1, result[1][i])
USER_NAME = getpass.getuser()
book.save(r"C:/Users/%s/Desktop/test.xls" % USER_NAME)