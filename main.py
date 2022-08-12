from glob import glob
from bs4 import BeautifulSoup
import xlsxwriter
import os

comprobantes = []


for filename in glob('comprobantes/*.htm'):
   with open(os.path.join(os.getcwd(), filename), 'r') as html_file:
        
        content = html_file.read()
        soup = BeautifulSoup(content, 'lxml').find(id="printArea_0").find("table").find("table")

        formTexts = map(lambda td: td.string.strip() if td.string else None, soup.find_all("td", class_="formText")[1:])
        formLabels = map(lambda td: td.string.strip() if td.string else None, soup.find_all("td", class_="formLabel"))

        comprobantes.append(dict(zip(formTexts, formLabels)))
        


workbook = xlsxwriter.Workbook('detalle_de_comprobantes.xlsx')
worksheet = workbook.add_worksheet()

for col_num, nombre_del_campo in enumerate(comprobantes[0].keys()):
    worksheet.write(0, col_num, nombre_del_campo)

for row_num, comprobante in enumerate(comprobantes):
    for col_num, valor_del_campo in enumerate(comprobante.values()):
        worksheet.write(row_num + 1, col_num, valor_del_campo)


workbook.close()