import re
from urllib.request import urlopen
from bs4 import BeautifulSoup
from xlwt import Workbook

wb = Workbook()
sheet1 = wb.add_sheet('Corredores')
sheet1.write(0, 0, 'Empresa')
sheet1.write(0, 1, 'Telefono')
sheet1.write(0, 2, 'Email')
sheet1.write(0, 3, 'Contacto')
sheet1.write(0, 4, 'Dirección')

link_base = 'https://www.portalinmobiliario.com'
quote_page = 'https://www.portalinmobiliario.com/empresas/corredoraspresentes.aspx'
page = urlopen(quote_page)
soup = BeautifulSoup(page, 'html.parser')
total_pages_text = soup.find('input', attrs={'name': 'ctl00$ContentPlaceHolder1$DataPager1$ctl02$ctl01'}).get('onclick')
total_pages = [s for s in total_pages_text if s.isdigit()]
total_pages = int("".join(total_pages))

enlaces = []
row = 0
for pagina in range(total_pages):
    pagina = str(pagina + 1)
    pagina = quote_page + f"?orden=Nombre&p={pagina}"
    print(pagina)
    pagina = urlopen(pagina)
    sopa = BeautifulSoup(pagina, 'html.parser')
    corredores_pagina = sopa.find('table', attrs={'id': 'ContentPlaceHolder1_ListViewCorredorasPresentes_groupPlaceholderContainer'}).find_all('tr')
    try:
        for corredor in corredores_pagina:
            if corredor.find_all('a'):
                for link in corredor.find_all('a'):
                    row += 1
                    link_corredor = link_base + link.get('href')
                    vista_corredor = urlopen(link_corredor)
                    soup_corredor = BeautifulSoup(vista_corredor, 'html.parser')
                    empresa = soup_corredor.find('h1').text
                    campos = soup_corredor.find_all('td', attrs={'class': 'Campo'})
                    valores = soup_corredor.find_all('td', attrs={'class': 'Valor'})
                    sheet1.write(row, 0, empresa)
                    for i, campo in enumerate(campos):
                        valor = re.sub(' +', ' ', valores[i].text)
                        if campo.text == 'Teléfonos':
                            sheet1.write(row, 1, valor)
                        elif campo.text == 'Email':
                            sheet1.write(row, 2, valor)
                        elif campo.text == 'Contacto':
                            sheet1.write(row, 3, valor)
                        elif campo.text == 'Dirección':
                            sheet1.write(row, 4, valor)
    except:
        pass

wb.save('corredores.xls')
