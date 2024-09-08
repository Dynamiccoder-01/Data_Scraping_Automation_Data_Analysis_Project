from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from bs4 import BeautifulSoup
from webdriver_manager.chrome import ChromeDriverManager
from io import StringIO


# Setup Selenium WebDriver with WebDriver Manager
chrome_options = Options()
# chrome_options.add_argument("--headless")  # Remove headless to see the browser window
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# Open the website
url = 'https://sdmdataaccess.nrcs.usda.gov/Query.aspx'
print("Current Window: ",url)
driver.get(url)

print("Website opened. Preparing to submit the query...")

try:
    # Wait for the query input field to be present
    query_input = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, 'TxtQuery'))  # Correct ID for the input field
    )
    submit_button = driver.find_element(By.ID, 'BtnSubmit')  # Correct ID for the submit button

    query = 'SELECT sacatalog.areaname AS "State", ROUND(SUM(chorizon.sandtotal_r * component.comppct_r / 100) / SUM(component.comppct_r), 2) AS "Sand Percentage" FROM sacatalog JOIN legend ON sacatalog.areasymbol = legend.areasymbol JOIN mapunit ON legend.lkey = mapunit.lkey JOIN component ON mapunit.mukey = component.mukey JOIN chorizon ON component.cokey = chorizon.cokey WHERE sacatalog.areaname IS NOT NULL GROUP BY sacatalog.areaname ORDER BY sacatalog.areaname;'

    query_input.clear()  # Clear any pre-existing text
    query_input.send_keys(query)
    submit_button.click()

    print("Query submitted. Waiting for results...")

    # Wait for the new window to open
    WebDriverWait(driver, 20).until(EC.new_window_is_opened)

    # Switch to the new window
    driver.switch_to.window(driver.window_handles[-1])

    # Print the current URL of the new window (for debugging)
    print(f"Switched to new window: {driver.current_url}")

    # Wait for the results table to be present (use generic table tag)
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.TAG_NAME, 'table'))  # Targeting the first table on the page
    )

    # Extract results from the page
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    table = soup.find('table')  # Find the first table in the HTML

    if table:
        print("Table found. Parsing data...")

        # Parse the table into a DataFrame
        html_string = str(table)
        df = pd.read_html(StringIO(html_string))[0]

        # Save the DataFrame to an Excel file
        excel_file = 'sand_percentage_report.xlsx'
        df.to_excel(excel_file, index=False, engine='openpyxl')

        print(f"Report successfully saved to {excel_file}")
    else:
        print("Error: Table not found in the response.")
finally:
    # Close the WebDriver
    driver.quit()
