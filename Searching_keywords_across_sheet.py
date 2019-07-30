import xlsxwriter
import pandas as pd
from pandas import DataFrame

class PrepareList():
    def __init__(self,app):
        self.app = app
        self.app_list = self.app.values.tolist()

        self.splitEachStringToList()
        self.removeSpaceInEachString()

    def splitEachStringToList(self):
        #split each string by (,)
        #"M,E" --> ['M','E']
        app_list2 = [] #temporary variable
        app_list3 = [] #temporary variable
        for i in range(len(self.app_list)):
            for j in range(len(self.app_list[i])):
                if isinstance(self.app_list[i][j], str):
                    if ',' in self.app_list[i][j]:
                        element = self.app_list[i][j].split(',')
                        app_list2.append(element)
                    else:
                        app_list2.append(self.app_list[i][j])
                else:
                    app_list2.append(self.app_list[i][j])
            app_list3.append(app_list2)
            app_list2 = [] #set app_list2 to default.
        self.app_list = app_list3

    def removeSpaceInEachString(self):
        #get rid of spaces in string
        #" ME" --> "ME"
        for i in range(len(self.app_list)):
            for j in range(len(self.app_list[i])):
                if isinstance(self.app_list[i][j], list):
                    for k in range(len(self.app_list[i][j])):
                        item_replace = self.app_list[i][j][k].replace(' ', '')
                        self.app_list[i][j][k] = item_replace

class ConfiguredOfficer():
    def __init__(self,com_list,app_list,colsname):
        self.com_list = com_list
        self.app_list = app_list
        self.conf_list = [] #A list for containing data for converting to dataframe
        self.grouped_list = [] #A list for result of grouping rows by server name
        self.columns = colsname
        self.conf_frame = pd.DataFrame(columns = self.columns)

        self.searchAndFillOfficerName()
        print(u'\u2713 ' + "Search and fill officer name.")
        self.groupByServerName()
        print(u'\u2713 ' + "Group rows by server names.")
        self.removeDuplicateItemInEachCell()
        print(u'\u2713 ' + "Remove duplicated officer names in each cell.")

    def searchAndFillOfficerName(self):
        configure_item = [] #A temporary list for containing founded keyword and wanted value
        for name in self.com_list:
            if isinstance(name,str): #name (or keyword) must be string ,not None //There are None in this list
                for i in range(len(self.app_list)):
                    for j in range(len(self.app_list[i])):
                        if isinstance(self.app_list[i][j],str) or isinstance(self.app_list[i][j], list): #We need only data type 'list' and 'string'. Others aren't cared.
                            if name in self.app_list[i][j]:
                                if isinstance(self.app_list[i][0],list):
                                    element = ','.join(self.app_list[i][0])
                                    self.app_list[i][0] = element
                                    element
                                if isinstance(self.app_list[i][1],list):
                                    element = ','.join(self.app_list[i][1])
                                    self.app_list[i][1] = element 
                                    element
                                configure_item.append(name + '|' + self.app_list[i][0] +'|'+ self.app_list[i][1])
                                break

        #Eliminate duplicated line
        configure_item = list(dict.fromkeys(configure_item))
        #sort item
        configure_item.sort()

        # "server1|header1|p1" --> ["server1","header","p1"]
        for element in configure_item:
            self.conf_list.append(element.split('|')) #split each string in configure_item by (|)
    
    def groupByServerName(self):
        #Convert list to dataframe
        self.conf_frame = pd.DataFrame(self.conf_list, columns = self.columns)
        #Group header_names and p_names by server_name
        grouped = self.conf_frame.groupby(['server_name'], sort=False).agg( ','.join)
        #reset index 
        grouped = grouped.reset_index()
        #Convert dataframe to list
        self.grouped_list = grouped.values.tolist()

    def removeDuplicateItemInEachCell(self):
        #Remove duplicate names in each cell
        temp = []
        for row in self.grouped_list:
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

    def getOfficerDataFrame(self):              
        return pd.DataFrame(self.grouped_list, columns = self.columns)

#-------------Main-------------#
conf_file = input()
out_path = input()

df = pd.ExcelFile('../'+conf_file) 
colsname = ['server_name' ,'header_name' ,'p_name']

com_sys = pd.read_excel(df, 'COMPUTERSYSTEM',usecols='G:I',header=None ,names=colsname ,skiprows=4)
app = pd.read_excel(df, 'APPLICATION' ,usecols='F:CN' ,skiprows=4)
print(u'\u2713 ' + "Initialize dataframe.")

com_list = com_sys.server_name.values.tolist()
app_list = PrepareList(app).app_list
print(u'\u2713 ' + "Prepare lists.")

configuredOfficer = ConfiguredOfficer(com_list,app_list,colsname)
conf_frame = configuredOfficer.getOfficerDataFrame()
print(u'\u2713 ' + "Convert configured list to dataframe.")

export_excel = conf_frame.to_excel('../'+out_path, index = None, header=True)
print(u'\u2713 ' + "Export configured dataframe to excel file.")
