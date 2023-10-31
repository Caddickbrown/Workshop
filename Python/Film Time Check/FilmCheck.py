import requests

from bs4 import BeautifulSoup

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
            film_name = li.strong.text.strip()
            # Check if ':' exists in the list item text
            if ':' in li.text:
                working_days = li.text.split(':')[1].strip()
            else:
                working_days = "Working days information not available."
            film_dict[film_name] = working_days
        for film, days in film_dict.items():
            print(f"{film}\n")
    else:
        print("Film list not found.")
else:
    print("Parent div not found.")
input("Press enter to exit ;)")