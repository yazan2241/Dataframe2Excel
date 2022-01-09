import pandas as pd
file = 'Zakat_Full_2021.xlsx'

xl = pd.ExcelFile(file)
sheet_names = xl.sheet_names

df_collection = {}
i = 0
for name in sheet_names :
    df_collection[i] = xl.parse(name)
    df_collection[i][df_collection[i].columns[0]] = df_collection[i][df_collection[i].columns[0]].fillna('NANN')
    i = i + 1

options = {}
options['strings_to_formulas'] = False
options['strings_to_urls'] = False
writer = pd.ExcelWriter('Zakat_Full_2022_mod11.xlsx',engine='xlsxwriter',options=options )

undef = 'NANN';
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
    worksheet.set_column('C:C',  16, format3)
    
    cnt=1
    i1 = 0
    j = 0
    elem = ''
    ddff = df_collection[i]
    
    for ii in range(ddff.shape[0]):
        for jj in range(1):
            value = ddff.iat[ii,jj]
            if value == undef :
                i1 = cnt
            else :
                if j != 0 :
                    if j == 2:
                        str1 = 'A'+str(j)
                    else :
                        str1 = 'A'+str(j+1)
                    str2 = 'A'+str(i1+1)
                    str3 = str1+':'+str2
                    worksheet.merge_range(str3,elem,header_format)
                elem = value
                j = cnt
                i1 = 0
        cnt = cnt + 1
        if cnt == ddff.shape[0] -1:
            str1 = 'A'+str(j+1)
            str2 = 'A'+str(i1+3)
            str3 = str1+':'+str2
            worksheet.merge_range(str3,elem,header_format)
        
writer.save()