import requests
import pandas as pd
import os
from openpyxl import load_workbook

os.chdir('C:/Users/Carlos_Lu/Desktop/BIOTOOLS/01_業務資料')

def get_table(url, year, pi):
    html = requests.get(url).content
    df_list = pd.read_html(html)
    df = df_list[-1]
    df = df[lambda x: x.年度 == year]
    df['PI'] = pi
    return df

All_project = pd.DataFrame({'PI':[],
                            '年度':[],
                            '核定經費(新台幣)':[],
                            '擔任工作':[],
                            '補助類別':[],
                            '學門代碼':[],
                            '計畫名稱':[],
                            })

#url = 'https://arspb.nstc.gov.tw/NSCWebFront/modules/talentSearch/talentSearch.do?action=initRsm17new&rsNo=eb3e61801dfa43f6b2640521233aef2d&LANG=chi'
#project_df = get_table(url, 113, '劉軒')
#All_project = pd.concat([All_project, get_table(url, 113, '邱亞芳')], axis=0, ignore_index=True)

All_project = pd.concat([All_project, get_table(input('Please enter URL: '), 113, input('Please enter PI name: '))], axis=0, ignore_index=True)

###Write projects to excel file
excel_workbook = load_workbook("113_Projects.xlsx")
writer = pd.ExcelWriter("113_Projects.xlsx", engine='openpyxl')

All_project.to_excel(writer, sheet_name = '中央大學', index = False)

writer.close()
