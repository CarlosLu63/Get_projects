from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time, os

os.chdir('C:/Users/Carlos_Lu/Desktop/BIOTOOLS/S_script/Get_projects')

# Set up the WebDriver (assuming you're using Chrome)
driver = webdriver.Chrome()

try:
    base_url = 'https://www.grb.gov.tw/search;keyword=undefined;type=GRB05'
    # Replace default keyword by the keyword I want to search.
    modified_url = base_url.replace('undefined', '微生物')

    # Get first results.
    driver.get(modified_url)
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
    df.to_excel('get_projects_113.xlsx', index=False)
    print('Data has been saved to output.xlsx')

finally:
    driver.quit()
    
    
