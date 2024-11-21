from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import time, os, requests, openpyxl
from lxml import html

os.chdir('C:/Users/Carlos_Lu/Desktop/BIOTOOLS/S_script/Get_projects')

def url_file_exist(url):
    response = requests.head(url)
    return {'exists': response.status_code == 200, 'status': response.status_code}

# Set up the WebDriver (assuming you're using Chrome)
driver = webdriver.Chrome()
driver.get("https://wsts.nstc.gov.tw/STSWeb/Award/AwardMultiQuery.aspx")

# Click on the "補助研究計畫" button
btn1 = driver.find_element(By.XPATH, "//*[@id='dtlItem_btnItem_0']")
btn1.click()
time.sleep(1)

# Click on the "生科處" option
unit_xpath2 = "//*[@id='wUctlAwardQueryPage_repQuery_ddlS1_5']/option[4]"
btn2 = driver.find_element(By.XPATH, unit_xpath2)
btn2.click()
time.sleep(1)

# Click on the "查詢" button
btn3 = driver.find_element(By.XPATH, '//*[@id="wUctlAwardQueryPage_btnQuery"]')
btn3.click()
time.sleep(1)

# Select "200筆"
btn4 = driver.find_element(By.XPATH, '//*[@id="wUctlAwardQueryPage_ddlPageSize"]/option[6]')
btn4.click()
time.sleep(5)

info = []
# Loop through pages
for j in range(10, 16):
    if j > 1:
        # Click on the "下一頁" button
        btn5 = driver.find_element(By.XPATH, '//*[@id="wUctlAwardQueryPage_grdResult_btnNext"]')
        btn5.click()
        time.sleep(5)

    page_source = driver.page_source
    tree = html.fromstring(page_source)

    # Loop through rows
    for i in range(2, 202):
        row_xpath = f'//*[@id="wUctlAwardQueryPage_grdResult"]/tbody/tr[{i}]'
        row = tree.xpath(row_xpath)
        if not row:
            continue

        row_text = row[0].text_content().strip()
        row_split = row_text.split('\t')
        row_split2 = row_split[3].split('\n')

        link_xpath = f'//*[@id="wUctlAwardQueryPage_grdResult"]/tbody/tr[{i}]//a'
        link = tree.xpath(link_xpath)
        if not link:
            continue

        link_href = link[0].get('onclick')
        link_href = link_href.split("no=")[1].split("', 'A")[0]

        # Get budget info
        link1 = f"https://wsts.nstc.gov.tw/STSWeb/Award/AwardDialog.aspx?year=112&sys=QS01&no={link_href}"
        if url_file_exist(link1)['exists']:
            link1_content = requests.get(link1).content.decode('utf-8')
            link1_text = html.fromstring(link1_content).text_content()
            link1_sub = link1_text.split("總核定金額：")[1].strip()
            link1_sub = link1_sub.replace("\r \r\t\r\n\n\n\r\n\r\n\r", "").replace("\r \r\n", ";")
            link1_split = link1_sub.split('\n')
        else:
            link1_split = ["", ""]

        # Get project summary
        link2 = f"https://wsts.nstc.gov.tw/STSWeb/Award/AwardDialog3.aspx?no={link_href}"
        if url_file_exist(link2)['exists']:
            link2_content = requests.get(link2).content.decode('utf-8')
            link2_text = html.fromstring(link2_content).text_content()
            link2_sub = link2_text.split("計畫概述：")[1].strip().replace('\r', '').replace('\t', '').replace('\n', '')
        else:
            link2_sub = ""

        info_row = [row_split[0], link_href, row_split[1], row_split[2], row_split2[0].replace("計畫名稱：", ""),
                    row_split2[2].replace("執行起迄：", ""), link1_split[0].replace("元", ""), link1_split[1], link2_sub]
        info.append(info_row)
        time.sleep(1)


driver.quit()






try:
    base_url = 'https://wsts.nstc.gov.tw/STSWeb/Award/AwardMultiQuery.aspx'

    # Get first results.
    driver.get(base_url)
    time.sleep(15)
        
    # Specify the range of the year I want to search.
    start_year_select = driver.find_element(By.CSS_SELECTOR, 'select[formcontrolname="planYearSt"')
    end_year_select = driver.find_element(By.CSS_SELECTOR, 'select[formcontrolname="planYearEn"')
    
    # Adjust listnumber in page
    pagenum_select = driver.find_element(By.CSS_SELECTOR, 'option[value="200"')
    pagenum_select.click()
    time.sleep(15)
    
    # Select the years. From... to...
    Select(start_year_select).select_by_value('113')
    Select(end_year_select).select_by_value('113')
    
    # Find search button and start to search.
    search_button = driver.find_element(By.CLASS_NAME, 'butsearch')
    search_button.click()
    time.sleep(15)

    # Fide page number
    max_page_number = int(len(driver.find_elements(By.CLASS_NAME, 'page')))
    
    # Extract data for 'conTitle' and 'conInfo'
    titles = []
    pi = []
    year = []
    expen = []

    if max_page_number == 0:
        con_titles = driver.find_elements(By.CLASS_NAME, 'conTitle')
        con_infos = driver.find_elements(By.CLASS_NAME, 'conInfo')
        
        assert len(con_titles) == len(con_infos), "Mismatch in number of titles and infos"
        
        for title, info in zip(con_titles, con_infos):
            titles.append(title.text)
            pi.append(info.text.split(' ')[1])
            year.append(info.text.split('：')[3].strip('當年度經費'))
            expen.append(info.text.split(' ')[3] + ' k')

    else:
        for page in range(1, max_page_number + 1):
            # Find all project title and project info
            con_titles = driver.find_elements(By.CLASS_NAME, 'conTitle')
            con_infos = driver.find_elements(By.CLASS_NAME, 'conInfo')
            
            # Ensure both lists are of the same length
            assert len(con_titles) == len(con_infos), "Mismatch in number of titles and infos"

            for title, info in zip(con_titles, con_infos):
                titles.append(title.text)
                pi.append(info.text.split(' ')[1])
                year.append(info.text.split('：')[3].strip('當年度經費'))
                expen.append(info.text.split(' ')[3] + ' k')
            
            # Navigate to next page and repeat finding project step
            if page < max_page_number:
                next_page_link = driver.find_element(By.LINK_TEXT, str(page + 1))
                next_page_link.click()
                # Wait for the new page to load
                time.sleep(5)
            
    # Create a DataFrame and save to Excel
    df = pd.DataFrame({
        'Project': titles,
        'PI': pi,
        'Year': year,
        'Expenditures': expen,
    })
    df.to_excel('../Get_project_results/get_projects_113.xlsx', index=False)
    print('Data has been saved to get_projects_113.xlsx')

finally:
    driver.quit()
    
    
