import xlsxwriter
import pandas as pd
from pandas import DataFrame


df = pd.ExcelFile("myfile.xlsx") 

colsname = ['server_name' ,'header_name' ,'p_name']

com_sys = pd.read_excel(df, 'COMPUTERSYSTEM',usecols='G:I',header=None ,names=colsname ,skiprows=4)
app = pd.read_excel(df, 'APPLICATION' ,usecols='F:CN' ,skiprows=4)


com_list = com_sys.server_name.values.tolist()
app_list = app.values.tolist()

#split each string by (,)
#"M,E" --> ['M','E']
app_list2 = [] #temporary variable
app_list3 = [] #temporary variable
for i in range(len(app_list)):
    for j in range(len(app_list[i])):
        if isinstance(app_list[i][j], str):
            if ',' in app_list[i][j]:
                element = app_list[i][j].split(',')
                app_list2.append(element)
            else:
                app_list2.append(app_list[i][j])
        else:
            app_list2.append(app_list[i][j])
    app_list3.append(app_list2)
    app_list2 = [] #set app_list2 to default.
app_list = app_list3

#get rid of spaces in string
#" ME" --> "ME"
for i in range(len(app_list)):
    for j in range(len(app_list[i])):
        if isinstance(app_list[i][j], list):
            for k in range(len(app_list[i][j])):
                item_replace = app_list[i][j][k].replace(' ', '')
                app_list[i][j][k] = item_replace


configure_item = [] #A list for containing founded keyword and wanted value
for name in com_list:
    if isinstance(name,str): #name (or keyword) must be string ,not None //There are None in this list
        for i in range(len(app_list)):
            for j in range(len(app_list[i])):
                if isinstance(app_list[i][j],str) or isinstance(app_list[i][j], list): #We need only data type 'list' and 'string'. Others aren't cared.
                    if name in app_list[i][j]:
                        if isinstance(app_list[i][0],list):
                            element = ','.join(app_list[i][0])
                            app_list[i][0] = element
                            element
                        if isinstance(app_list[i][1],list):
                            element = ','.join(app_list[i][1])
                            app_list[i][1] = element 
                            element
                        configure_item.append(name + '|' + app_list[i][0] +'|'+ app_list[i][1])
                        break

#Eliminate duplicated line
configure_item = list(dict.fromkeys(configure_item))
#sort item
configure_item.sort()

# "server1|header1|p1" --> ["server1","header","p1"]
conf_list = [] #A list for containing data for convert to dataframe
for element in configure_item:
    conf_list.append(element.split('|')) #split each string in configure_item by (|)

#Convert list to dataframe
conf_frame = pd.DataFrame(conf_list, columns = ['server_name','header_name','p_name'])

#Group header_names and p_names by server_name
grouped = conf_frame.groupby(['server_name'], sort=False).agg( ','.join)
#reset index 
grouped = grouped.reset_index()
#Convert dataframe to list
grouped_list = grouped.values.tolist()

#Remove duplicate names in each cell
temp = []
for row in grouped_list:
    if ',' in row[0]:
        temp = row[0].split(',')
        temp = list(dict.fromkeys(temp))
        temp = ','.join(temp)
        row[0] = temp

    if ',' in row[1]:
        temp = row[1].split(',')
        temp = list(dict.fromkeys(temp))
        temp = ','.join(temp)
        row[1] = temp
        
    if ',' in row[2]:
        temp = row[2].split(',')
        temp = list(dict.fromkeys(temp))
        temp = ','.join(temp)
        row[2] = temp
                
#Convert list to dataframe
conf_frame = pd.DataFrame(grouped_list, columns = ['server_name','header_name','p_name'])

#Export to .xlsx
export_excel = conf_frame.to_excel(r'C:\mypath\myfile.xlsx', index = None, header=True)