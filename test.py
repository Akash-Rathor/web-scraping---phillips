from bs4 import BeautifulSoup as bs
import requests
import pandas as pd
import pprint

context = {

    'SKU Name':[],
    'SKU code':[],
    'MRP':[],
    'Photographs':[],
    'All technical Specification':[],
}
headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Max-Age': '3600',
    'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/52.0'
}

def get_product_data(url):

    r = requests.get(url, headers=headers)
    page = bs(r.content, 'xml')

    section = page.find_all('div',{'class':'p-pc05v2__card--layout'})

    for div in section:
        name= div.find('span',{'class':'p-heading-bold'}).string+" "+div.find('span',{'class':'p-heading-light'}).string
        if name !=None:
            context['SKU Name'].append(name)
            try:
                context['SKU code'].append(div.find('p',{'class':'p-pc05v2__card-ctn p-body-copy-03 p-heading-light'}).string)
            except:
                context['SKU code'].append("")
            try:
                mp = div.find('span',{'class':'p-current-price-value'})
                context['MRP'].append(mp)
                
            except:
                context['MRP'].append("")
            try:
                context['Photographs'].append(div.find('img').get('src'))
            except:
                context['Photographs'].append("")
            try:
                tech = div.find('ul',{'class':'p-bullets p-heading-light'}).find_all('li')
                p = []
                for i in tech:
                    p.append(i.string)
                context['All technical Specification'].append(",".join(p))
            except:
                context['All technical Specification'].append("")
    

def get_page(url):

    headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Max-Age': '3600',
    'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/52.0'
    }

    r = requests.get(url, headers=headers)
    page = bs(r.content, 'xml')
    try:
        urls = page.find('span',{'class':'p-d06__total-pages'}).string
    except:
        urls=0
    return int(urls)

def make_url():
    df = pd.read_excel('New.xlsx',sheet_name="Product Catagories")
    for i,j in zip(df['Category'],df['Start']):

        url = "https://www.philips.co.in/"+"-".join(j.split(" ")).replace("--","-")+'/'+"-".join(i.split(" ")).replace("--","-")+"/all#"

        urls = get_page(url.lower())
        
        get_product_data(url.lower())

        for i in range(urls):
            l = (url.lower()+'page={num}&layout=12').format(num=i+1)
            print(l)
            get_product_data(l)
            
        print("one url completed")

make_url()


# To Store Data and format in excel
xs = pd.DataFrame.from_dict(context, orient='index').transpose()
dfsn = [(xs,'Product Catagories'),]
writer = pd.ExcelWriter('Test_Output.xlsx', engine = 'xlsxwriter')
workbook  = writer.book

for i in range(len(dfsn)):
    dfsn[i][0].to_excel(writer,sheet_name = dfsn[i][1],index=False)
    format1  = workbook.add_format({'font_color':'#9c0005','bg_color':'#ffc7ce'}) #red
    format2  =  workbook.add_format({'align': 'center','font_size':11,'font_name':'Calibri','font_color':'#006100','bg_color':'#c6efce','border': 2,'border_color':'#f2665f'})
    format3  =  workbook.add_format({'align': 'center','font_size':11,'font_name':'Calibri','font_color':'#006100','bg_color':'#a8fff6','border': 2,'border_color':'#f2665f'})
    worksheet = writer.sheets[dfsn[i][1]]

    header_format = workbook.add_format({
    'bold': True,
    'font_name':'Century Gothic',
    'border': 2,
    'bg_color':'#302e45',
    'font_size':10,
    'font_color':'#fdb515',
    'align':'center'})

    worksheet.set_column("A2:E"+str(len(dfsn[i][0])),None,format2)

    worksheet.conditional_format("A2:E"+str(len(dfsn[i][0])+1), {'type':'blanks',
    'format':format3})

    for col_num, value in enumerate(dfsn[i][0].columns.values):
        worksheet.write(0, col_num, value, header_format)

        column_len = dfsn[i][0][value].astype(str).str.len().max()
        column_len = max(column_len, len(value)) + 2
        worksheet.set_column(col_num, col_num, column_len)
writer.save()
