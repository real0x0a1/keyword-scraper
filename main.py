#!/bin/python3

# -*- Author: real0x0a1 (Ali) -*-


import urllib3
import xlwt

from rich.console import Console
from rich.progress import Progress
from rich.prompt import Prompt
from bs4 import BeautifulSoup

# Initialize console for rich text output
console = Console()

# Prompt user to enter a keyword, defaulting to "python3" if no input is provided
keyword_query = Prompt.ask("Enter the keyword to search:", default="python3")

# List to store extracted keywords
extracted_keywords = []

# Create a session to manage requests
http_manager = urllib3.PoolManager()
response = http_manager.request('GET', f'https://www.google.com/search?q={keyword_query}')
parsed_content = BeautifulSoup(response.data, 'html.parser')

# Extract keywords from the parsed HTML content
for element in parsed_content.find_all('div', {'class': 'BNeawe s3v9rd AP7Wnd lRVwie'}):
    extracted_keywords.append(element.text)

# Create a new Excel workbook and add a worksheet
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('Keyword Data')

# Progress bar to monitor the data writing process
with Progress() as progress:
    data_write_task = progress.add_task("Writing data to Excel...", total=len(extracted_keywords))

    # Write each keyword and its related data to the Excel sheet
    for index, keyword in enumerate(extracted_keywords):
        keyword_response = http_manager.request('GET', f'https://www.google.com/search?q={keyword}')
        keyword_content = BeautifulSoup(keyword_response.data, 'html.parser')

        worksheet.write(0, index, keyword)
        related_items = keyword_content.find_all('div', {'class': 'BNeawe s3v9rd AP7Wnd lRVwie'})

        for row, item in enumerate(related_items):
            worksheet.write(row + 1, index, item.text)

        progress.update(data_write_task, advance=1)

# Notify user of success and save the Excel file
console.print('Success, view the [green]keywords.xls[/green] file.')
workbook.save(f'{keyword_query}.xls')
