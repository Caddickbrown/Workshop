import requests
from bs4 import BeautifulSoup
from tabulate import tabulate

url = 'https://analoguewonderland.co.uk/pages/wonderlab-status'
response = requests.get(url)
content = response.text

soup = BeautifulSoup(content, 'html.parser')

parent_div = soup.find('div', class_='page__content')

if parent_div:
    film_list = parent_div.find('ul')
    if film_list:
        film_dict = {}
        for li in film_list.find_all('li'):
            text_parts = li.text.split(':')
            if len(text_parts) >= 2:
                film_name = text_parts[0].strip()
                working_days = text_parts[1].strip()
            else:
                film_name = li.text.strip()
                working_days = "Working days information not available."
            
            # Check for the stop condition
            if "Note - Working days are Monday to Friday, and we are closed on Bank Holidays" in film_name:
                break
            
            film_dict[film_name] = working_days
        
        table_data = []
        for film, days in list(film_dict.items()):
            table_data.append([film, days])
        print((tabulate(table_data, headers=["Film", "Working Days"], tablefmt="fancy_grid")))
    else:
        print("Film list not found.")
else:
    print("Parent div not found.")