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
import openpyxl
from openpyxl.drawing.image import Image

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
                      ROUND(SUM(chorizon.sandtotal_r * component.comppct_r / 100) / SUM(component.comppct_r), 2) AS "Sand Percentage" 
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

        # Define a function to extract city from state
        def extract_city(state):
            return state.split(',')[0].strip()

        # Apply the function to create the 'City' column
        df['City'] = df['State'].apply(extract_city)

        # Aggregate data by City, averaging the Sand Percentage
        df_aggregated = df.groupby('City').agg({'Sand Percentage': 'mean'}).reset_index()

        # Sort by 'Sand Percentage' and select top 30
        df_top30 = df_aggregated.sort_values(by='Sand Percentage', ascending=False).head(30)

        # Save both DataFrames to an Excel file
        excel_file = 'sand_percentage_report.xlsx'
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            # Save original data
            df.to_excel(writer, sheet_name='Original Data', index=False)

            # Save top 30 aggregated data
            df_top30.to_excel(writer, sheet_name='Top 30 Cities', index=False)

        print(f"Reports successfully saved to {excel_file}")

        # Create a vertical bar plot with different colors
        plt.figure(figsize=(16, 10))
        # Create a hue column with dummy values for color palette
        df_top30['Hue'] = df_top30.index
        bar_plot = sns.barplot(x='City', y='Sand Percentage', data=df_top30,
                               palette=sns.color_palette("husl", len(df_top30)), hue='Hue', dodge=False)
        plt.xticks(rotation=90)  # Rotate city names for clarity
        plt.title('Top 30 Cities by Sand Percentage')
        plt.tight_layout()

        # Add values on top of bars with bold font
        for p in bar_plot.patches:
            height = p.get_height()
            bar_plot.annotate(f'{height:.2f}',
                              (p.get_x() + p.get_width() / 2., height),
                              ha='center', va='center',
                              xytext=(0, 5),
                              textcoords='offset points',
                              fontsize=12,
                              fontweight='bold')

        # Save the plot to a PNG file
        plot_file = 'top30_sand_percentage_barplot.png'
        plt.savefig(plot_file)
        plt.close()

        print(f"Bar plot saved as {plot_file}")

        # Add the plot to the Excel file
        with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
            workbook = writer.book
            worksheet = workbook.create_sheet('Bar Plot')

            # Add the plot image to the Excel file
            img = Image(plot_file)
            worksheet.add_image(img, 'A1')

        print("Bar plot with different colors and bold values added to the Excel file.")
    else:
        print("Error: Table not found in the response.")

finally:
    # Close the WebDriver
    driver.quit()
