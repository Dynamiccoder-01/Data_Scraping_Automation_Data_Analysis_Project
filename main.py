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
import matplotlib.pyplot as plt
import seaborn as sns

# Setup Selenium WebDriver with WebDriver Manager
chrome_options = Options()
# chrome_options.add_argument("--headless")  # Remove headless to see the browser window
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# Open the website
url = 'https://sdmdataaccess.nrcs.usda.gov/Query.aspx'
print("Current Window: ", url)
driver.get(url)

print("Website opened. Preparing to submit the query...")

try:
    # Wait for the query input field to be present
    query_input = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, 'TxtQuery'))  # Correct ID for the input field
    )
    submit_button = driver.find_element(By.ID, 'BtnSubmit')  # Correct ID for the submit button

    query = '''SELECT sacatalog.areaname AS "State", 
                      ROUND(SUM(chorizon.sandtotal_r * component.comppct_r / 100) / SUM(component.comppct_r), 2) 
                      AS "Sand Percentage" 
               FROM sacatalog 
               JOIN legend ON sacatalog.areasymbol = legend.areasymbol 
               JOIN mapunit ON legend.lkey = mapunit.lkey 
               JOIN component ON mapunit.mukey = component.mukey 
               JOIN chorizon ON component.cokey = chorizon.cokey 
               WHERE sacatalog.areaname IS NOT NULL 
               GROUP BY sacatalog.areaname 
               ORDER BY sacatalog.areaname;'''

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

        # Save the DataFrame to an Excel file with xlsxwriter engine
        excel_file = 'sand_percentage_report_with_graph.xlsx'
        writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Sand Data')

        # Create a plot using Seaborn
        plt.figure(figsize=(12, 8))  # Adjust figure size based on dataset size
        sns.lineplot(data=df, x='State', y='Sand Percentage', marker='o')

        # Improve plot readability
        plt.xticks(rotation=90)  # Rotate X-axis labels for better readability
        plt.title('Sand Percentage by State', fontsize=16)
        plt.xlabel('State', fontsize=12)
        plt.ylabel('Sand Percentage', fontsize=12)
        plt.grid(True)  # Add gridlines for better clarity
        plt.tight_layout()  # Ensure labels and title fit within the figure

        # Save the plot as an image file
        plot_image = 'sand_percentage_plot.png'
        plt.savefig(plot_image)

        print(f"Plot saved as {plot_image}")

        # Access the workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Sand Data']

        # Insert the plot image into the worksheet
        worksheet.insert_image('E2', plot_image)

        # Save and close the Excel writer
        writer.close()

        print(f"Report with graph successfully saved to {excel_file}")

    else:
        print("Error: Table not found in the response.")

finally:
    # Close the WebDriver
    driver.quit()
