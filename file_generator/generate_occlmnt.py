import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import argparse

DATA_SOURCE = 'Horizons;GaiaEDR3' 

def fetch_event_list(object_id, date):
    url = f"https://www.occultwatcher.net/api2/v1/owc/events-all/?objectId={object_id}&dt={date}&bf=1&t=false"
    response = requests.get(url)
    data = response.json()
    return data

def create_webdriver():
    service = Service('/opt/homebrew/bin/chromedriver')  # Adjust the path if necessary
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def fetch_occlmnt_data(occultation_id):
    url = f"https://cloud.occultwatcher.net/event/{occultation_id}/{DATA_SOURCE}"

    driver = create_webdriver()

    try:
        # Fetch the page with Selenium
        driver.get(url)

        # Increase the wait time to ensure the page loads completely
        WebDriverWait(driver, 30).until(
            EC.text_to_be_present_in_element((By.TAG_NAME, 'pre'), 'occelmnt file')
        )

        # Get the page source and parse it with BeautifulSoup
        html_content = driver.page_source
        soup = BeautifulSoup(html_content, 'html.parser')

        # Locate the specific `<pre>` tag containing the occultation element file
        pre_tags = soup.find_all('pre', class_='mt-3')
        for tag in pre_tags:
            text = tag.get_text()
            if "Occult  BEGIN" in text and "Occult  END" in text:
                return text

    finally:
        driver.quit()

    # If not found, return None
    return None

def main():
    parser = argparse.ArgumentParser(description="Fetch occultation data for a given object ID and date.")
    parser.add_argument("object_id", type=int, help="The object ID to fetch data for.")
    parser.add_argument("date", type=str, help="The date to fetch data for (YYYY-MM-DD).")
    
    args = parser.parse_args()
    
    object_id = args.object_id
    date = args.date
    data = fetch_event_list(object_id, date)

    if len(data) > 1:
        raise ValueError("More than one event found for the given date and object ID")
    
    occultation_id = data[0]['id']
    occultation_data = fetch_occlmnt_data(occultation_id)
    print(occultation_data)

if __name__ == "__main__":
    main()
