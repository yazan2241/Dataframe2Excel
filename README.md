# Dataframe2Excel
convert excel to dataframe and reconstruct new excel file with handling merge cells

# Contents

1 - col.py
2 - row.py


# Goal

the goal of this repo is to get in touch with merged cells and merged columns in a multisheet Excel file
how to handle it and how to re construct a merged cells when creating a new excel file

I had an excel file which has some merged cells in rows and in columns , so i wanted to make some changes to the file using python and re construct the excel file after i finish


# Read Excel
I used pandas read_excel libraby to read the file

Then created a collection List which store each sheet seperatly in a new dataframe

for merged columns , pandas dataframe stores the column name as 'Unnamed-i'
for merged rows , pandas dataframe stores the row name as empty string ''


# Handle merged cell

col.py : file which handle merged columns
row.py : file which handle merged rows




