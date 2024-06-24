import requests
from bs4 import BeautifulSoup

url = 'https://analoguewonderland.co.uk/pages/wonderlab-status'
response = requests.get(url)
content = response.text
soup = BeautifulSoup(content, 'html.parser')

parent_div = soup.find('div', class_='page__content')

if parent_div:
    film_list = parent_div.find('ul')
    if (film_list):
        film_dict = {}
        for li in film_list.find_all('li'):
            li_text = li.text.strip()
            
            # Check for the stop condition
            if "Note - Working days are Monday to Friday, and we are closed on Bank Holidays" in li_text:
                break
            
            text_parts = li_text.split(':')
            if len(text_parts) >= 2:
                film_name = text_parts[0].strip()
                working_days = text_parts[1].strip()
            else:
                film_name = li_text
                working_days = "Working days information not available."
            
            film_dict[film_name] = working_days
        
        print("Film                     | Working Days")
        print("----------------------------------------")
        for film, days in film_dict.items():
            print(f"{film:<25} | {days}")
    else:
        print("Film list not found.")
else:
    print("Parent div not found.")
