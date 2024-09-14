# Data Scraping and Visualization Project

This project automates the extraction of soil sand percentage data from the NRCS SDM Data Access web portal. It processes the scraped data to generate visual representations and stores them in an Excel file, making it easy to analyze and interpret.

## Table of Contents
- [Features](#features)
- [Technologies Used](#technologies-used)
- [Installation](#installation)
- [Usage](#usage)
- [Example Output](#example-output)
- [Contributing](#contributing)
- [License](#license)

## Features
- **Automated Web Scraping**: Uses Selenium to automate the process of submitting queries and retrieving data from the NRCS SDM portal.
- **Data Parsing and Processing**: Extracted data is parsed and cleaned using BeautifulSoup and Pandas.
- **City-Level Data Analysis**: A new column for city names is created based on the state data. Duplicate city entries are averaged.
- **Data Visualization**: A vertical bar chart is generated, displaying the top 30 cities by sand percentage, with colorful bars and bold labels for clarity.
- **Excel Reporting**: The scraped data and the chart are saved in an Excel file with multiple sheets, including the data and visualizations.

## Technologies Used
- **Python**
- **Selenium**: For browser automation.
- **BeautifulSoup**: For parsing HTML content.
- **Pandas**: For data manipulation and analysis.
- **Seaborn & Matplotlib**: For data visualization.
- **OpenPyXL**: For exporting data and charts into Excel.
- **WebDriver Manager**: For managing the browser driver automatically.

## Installation
1. **Clone the repository**:
   ```bash
   git clone https://github.com/your-username/repository-name.git
2. **Navigate to the project directory:**:
   ```bash
   cd repository-name
3. **Install dependencies:**:
   ```bash
   pip install selenium beautifulsoup4 pandas seaborn openpyxl matplotlib webdriver-manager
4. **Set up ChromeDriver:**:
   -You can download it manually or let WebDriver Manager handle it automatically.

## Usage

1. **Run the script:**:
   ```bash
   python main.py
2. **After execution, an Excel file named sand_percentage_report.xlsx will be generated. It contains:**:
   -A sheet with the raw data (State, City, Sand Percentage).
   -A sheet with a vertical bar chart visualizing the top 30 cities with the highest sand percentage values.
