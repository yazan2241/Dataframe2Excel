import pandas as pd

file = 'Zakat_Full_2021.xlsx'

df = pd.read_excel(file)

xl = pd.ExcelFile(file)
sheet_names = xl.sheet_names
sheet_names

df_collection = {}
i = 0
for name in sheet_names :
    df_collection[i] = xl.parse(name)
    df_collection[i][df_collection[i].columns[0]] = df_collection[i][df_collection[i].columns[0]].fillna(method='ffill', axis=0)
    i = i + 1


i = 0
name = '';
undef = 'Unnamed';
for name in sheet_names :
    for col in df_collection[i] .columns:
        if undef in col :
            df_collection[i] = df_collection[i].rename(columns={col: name})
        else :
            name = col
    i = i + 1


def getDate(name) :
    if name == "January  2021" :
        return "1/2021"
    if name == "February  2021" :
        return "2/2021"
    if name == "March  2021" :
        return "3/2021"
    if name == "First Quarter  2021" :
        return "Q1/2021"
    if name == "April  2021" :
        return "4/2021"
    if name == "May  2021" :
        return "5/2021"
    if name == "Jun 2021" :
        return "6/2021"
    if name == "July 2021" :
        return "7/2021"
    if name == "Second Quarter 2021" :
        return "Q2/2021"
    if name == "August 2021" :
        return "8/2021"
    if name == "Sep 2021" :
        return "9/2021"
    if name == "Oct 2021" :
        return "10/2021"
    if name == "Oct 2021" :
        return "10/2021"
    if name == "Nov 2021" :
        return "11/2021"
    if name == "Dec 2021" :
        return "12/2021"
    if name == "Third Quarter 2021" :
        return "Q3/2021"


i=0
for name in sheet_names :
    df_collection[i].insert(0, 'التاريخ', getDate(name))
    i = i + 1



options = {}
options['strings_to_formulas'] = False
options['strings_to_urls'] = False
writer = pd.ExcelWriter('Zakat_Full_2022_mod4.xlsx',engine='xlsxwriter',options=options )


for i in range(8) :    
    df_collection[i].to_excel(writer, sheet_names[i] , index=False)
    workbook  = writer.book
    worksheet = writer.sheets[sheet_names[i]]
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'vcenter',
        'align': 'center',
        'bg_color': '#305496',
        'font_color': '#ffffff'
    })
    format2 = workbook.add_format({'bg_color': '#8EA9DB','font_color': '#ffffff' , 'align': 'center'})
    format3 = workbook.add_format({'bg_color': '#D9E1F2' })
    format4 = workbook.add_format({'bg_color': '#37658F','font_color': '#ffffff' , 'align': 'center'})
    worksheet.set_column('A:A',  16, format4)
    worksheet.set_column('B:B',  16, format2)
    worksheet.set_column('C:C',  16, format3)
    
    for col_num, value in enumerate(df_collection[i].columns.values):
        worksheet.write(0, col_num, value, header_format)

writer.save()