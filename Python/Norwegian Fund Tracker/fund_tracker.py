import requests
from bs4 import BeautifulSoup
import time

def get_fund_tracker_value():
  url = "https://www.nbim.no/en/"
  headers = {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
  }  # Adding a user agent to prevent 403 Forbidden error
  response = requests.get(url, headers=headers)
  if response.status_code == 200:
    soup = BeautifulSoup(response.text, "html.parser")
    # Extracting the fund tracker value using the id "liveNavNumber"
    fund_tracker_value = soup.find("span", id="liveNavNumber")
    if fund_tracker_value:
      return fund_tracker_value.text.strip()
    else:
      print("Could not find the fund tracker value.")
      return None
  else:
    print("Failed to fetch data from the website.")
    return None

def get_previous_value():
  # Store the previously retrieved value in a variable
  previous_value = None
  return previous_value

def compare_values(current_value, previous_value):
  # Check if the current value is greater than the previous value
  if current_value > previous_value:
    return "up"
  # Check if the current value is less than the previous value
  elif current_value < previous_value:
    return "down"
  # Otherwise, the values are the same
  else:
    return "unchanged"

while True:
  current_value = get_fund_tracker_value()
  if current_value:
    previous_value = get_previous_value()
    # Update the previously retrieved value
    get_previous_value.update(current_value)
    direction = compare_values(current_value, previous_value)
    print(f"Fund Tracker Value: {current_value} ({direction})")
  else:
    print("No data available.")
  time.sleep(5)
