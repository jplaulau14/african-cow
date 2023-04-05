import os
import sqlite3
import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By


def download_excel(url: str, css_selector: str) -> None:
    """
    Downloads the excel data file from the given URL and saves it to the current directory.
    
    Parameters:
        url (str): The URL to the data file.
        css_selector (str): The CSS selector for the download button.

    Returns:    
        None
    """
    s = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=s)
    driver.get(url)

    # Locate the download button and get the CSV file URL
    download_button = driver.find_element(By.CSS_SELECTOR, '.wp-block-button__link')
    csv_url = download_button.get_attribute('href')

    # Download the CSV file
    response = requests.get(csv_url)
    with open('skill_test_data.xlsx', 'wb') as f:
        f.write(response.content)

    # Close the webdriver
    driver.quit()


def create_pivot_table(data: pd.DataFrame) -> pd.DataFrame:
    """
    Creates a pivot table from the given data and saves it to a new Excel file.

    Parameters:
        data (DataFrame): The data to create the pivot table from.
    
    Returns:
        DataFrame: The pivot table.
    """
    pivot_table = pd.pivot_table(data, index='Platform (Northbeam)',
                                values=['Spend', 'Attributed Rev (1d)', 'Imprs',
                                        'Visits', 'New Visits', 'Transactions (1d)',
                                        'Email Signups (1d)'],
                                aggfunc='sum')

    column_sort = ['Spend', 'Attributed Rev (1d)', 'Imprs',
                                        'Visits', 'New Visits', 'Transactions (1d)',
                                        'Email Signups (1d)']

    # Sort the columns
    pivot_table = pivot_table.reindex(column_sort, axis=1)

    # Rename the column headers
    pivot_table.columns = ['Sum of Spend', 'Sum of Attributed Rev (1d)', 'Sum of Imprs',
                        'Sum of Visits', 'Sum of New Visits', 'Sum of Transactions (1d)',
                        'Sum of Email Signups (1d)']

    # Sort by revenue descending
    pivot_table = pivot_table.sort_values('Sum of Attributed Rev (1d)', ascending=False)

    for column in pivot_table.columns:
        pivot_table[column] = pivot_table[column].apply(lambda x: f"{x:,.2f}")

    # Apply the formatting for the "Sum of Spend" and "Sum of Attributed Rev (1d)" columns
    pivot_table['Sum of Spend'] = pivot_table['Sum of Spend'].apply(lambda x: f"${x}")
    pivot_table['Sum of Attributed Rev (1d)'] = pivot_table['Sum of Attributed Rev (1d)'].apply(lambda x: f"${x}")

    return pivot_table


def create_database_and_insert_data(database_name: str, table_name: str, data: pd.DataFrame) -> None:
    """
    Creates a new SQLite database and inserts the given data into a new table.

    Parameters:
        database_name (str): The name of the database.
        table_name (str): The name of the table.
        data (DataFrame): The data to insert into the table.

    Returns:
        None
    """
    conn = sqlite3.connect(database_name)
    data.to_sql(table_name, conn, if_exists='replace')
    conn.commit()
    conn.close()


if __name__ == "__main__":
    # Step 1: Download the data file
    url = "https://jobs.homesteadstudio.co/data-engineer/assessment/download/"
    css_selector = ".wp-block-button__link'"
    download_excel(url, css_selector)

    # Step 2: Create a pivot table
    input_file = "skill_test_data.xlsx"
    data = pd.read_excel(input_file, sheet_name="data")
    pivot_table = create_pivot_table(data)

    # Save the pivot table to a new Excel file
    pivot_table.to_excel("pivot_table.xlsx")

    # Step 3 and 4: Create a new SQLite database and insert the pivot table
    database_name = "african_cow.db"
    table_name = "pivot_table"
    create_database_and_insert_data(database_name, table_name, pivot_table)

    print("Process completed.")