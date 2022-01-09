import pandas as pd
file = 'Zakat_Full_2021.xlsx'

xl = pd.ExcelFile(file)
sheet_names = xl.sheet_names

i = 0
name = '';
df_collection = {}
undef = 'Unnamed';
for name in sheet_names :
    df_collection[i] = xl.parse(name)
    for col in df_collection[i] .columns:
        if col == 'nan' :
            df_collection[i] = df_collection[i].rename(columns={col: 'Unnamed'})
        else :
            name = col
    i = i + 1

options = {}
options['strings_to_formulas'] = False
options['strings_to_urls'] = False
writer = pd.ExcelWriter('Zakat_Full_2022_mod97.xlsx',engine='xlsxwriter',options=options )

def getChar(num):
    if num == 0 :
        return 'A'
    if num == 1 :
        return 'B'
    if num == 2 :
        return 'C'
    if num == 3 :
        return 'D'
    if num == 4 :
        return 'E'
    if num == 5 :
        return 'F'
    if num == 6 :
        return 'G'
    if num == 7 :
        return 'H'
    if num == 8 :
        return 'I'
    if num == 9 :
        return 'J'
    if num == 10 :
        return 'K'
    if num == 11 :
        return 'L'
    if num == 12 :
        return 'M'
    if num == 13 :
        return 'N'
    if num == 14 :
        return 'O'
    if num == 15 :
        return 'P'
    if num == 16 :
        return 'Q'
    if num == 17 :
        return 'R'
    if num == 18 :
        return 'S'
    if num == 19 :
        return 'T'


undef = 'Unnamed';
workbook  = writer.book
for i in range(8) : 
    df_collection[i].to_excel(writer, sheet_names[i] , index=False)
    worksheet = writer.sheets[sheet_names[i]]
    
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'vcenter',
        'align': 'center',
        'bg_color': '#305496',
        'font_color': '#ffffff'
    })
    format2 = workbook.add_format({'bg_color': '#D9E1F2','font_color': '#ffffff' , 'align': 'center'})
    format3 = workbook.add_format({'bg_color': '#D9E1F2' , 'align': 'center'})
    format4 = workbook.add_format({'bg_color': '#8EA9DB','font_color': '#ffffff' , 'align': 'center'})
    worksheet.set_column('A:A',  16, format4)
    worksheet.set_column('B:B',  16, format2)
    
    
    cnt=1
    i1 = 0
    j = 0
    elem = ''
    ddff = df_collection[i]
    
    for col in df_collection[i] .columns:
        if undef in col :
            i1 = cnt
        else :
            if j !=0 :
                str1 = getChar(j-1)+'1'
                str2 = getChar(i1-1)+'1'
                str3 = str1+':'+str2
                if j == i1 :
                    worksheet.write(0, j-1, elem, header_format)
                else :
                    worksheet.merge_range(str3,elem,header_format)
            elem = col
            j = cnt
            i1 = cnt
        cnt = cnt + 1
        if cnt == df_collection[i].shape[1] -1:
            str1 = getChar(j-1)+'1'
            str2 = getChar(i1+1)+'1'
            str3 = str1+':'+str2
            worksheet.merge_range(str3,elem,header_format)
        
writer.save()