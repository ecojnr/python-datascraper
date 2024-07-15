import requests
from bs4 import BeautifulSoup
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
import openpyxl
import time
import os
import ssl

# Custom HTTPSAdapter to disable SSL verification
class TLSAdapter(requests.adapters.HTTPAdapter):
    def init_poolmanager(self, *args, **kwargs):
        context = ssl.create_default_context()
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
        kwargs['ssl_context'] = context
        return super().init_poolmanager(*args, **kwargs)

    def proxy_manager_for(self, *args, **kwargs):
        context = ssl.create_default_context()
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
        kwargs['ssl_context'] = context
        return super().proxy_manager_for(*args, **kwargs)

def search_google(query):
    edge_driver_path = r"C:\Users\thoma\Desktop\msedgedriver.exe"  # Ensure this path is correct
    user_data_dir = r"C:\Users\thoma\AppData\Local\Microsoft\Edge\User Data"  # Adjust this to your actual Edge user data path
    
    options = webdriver.EdgeOptions()
    options.use_chromium = True
    options.add_argument(f"user-data-dir={user_data_dir}")
    
    # Ensure only one instance of Edge is running
    os.system("taskkill /F /IM msedge.exe")
    time.sleep(2)

    service = Service(executable_path=edge_driver_path)
    driver = webdriver.Edge(service=service, options=options)
    
    driver.get('https://www.google.com/')
    wait = WebDriverWait(driver, 30)

    try:
        # Handle Google's cookie consent
        consent_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[text()="I agree" or text()="Accept all"]')))
        if consent_button:
            print("Found consent button, clicking it.")
            consent_button.click()
            time.sleep(2)  # Give it a moment to process the click
    except Exception as e:
        print("Consent button not found or another issue occurred:", e)

    try:
        # Wait for the search box to be visible and interactable
        search_box = wait.until(EC.visibility_of_element_located((By.NAME, 'q')))
        print("Search box found, entering query.")
        search_box.send_keys(query)
        search_box.send_keys(Keys.RETURN)
    except Exception as e:
        print("Search box not interactable or another issue occurred:", e)
        driver.quit()
        raise

    try:
        # Print the page source for debugging
        print("Page source before waiting for search results:\n", driver.page_source)

        # Wait for the search results to be present
        print("Waiting for search results...")
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div#search a')))
        print("Search results found.")
        
        # Print the page source after waiting for search results for further debugging
        print("Page source after waiting for search results:\n", driver.page_source)
    except Exception as e:
        print("Search results not found within the time limit:", e)
        print(driver.page_source)  # Print page source for debugging
        driver.quit()
        raise

    links = []
    results = driver.find_elements(By.CSS_SELECTOR, 'div#search a')
    for result in results:
        link = result.get_attribute('href')
        links.append(link)
    
    driver.quit()
    return links

def extract_info(url):
    try:
        session = requests.Session()
        # Mount the custom adapter to disable SSL verification
        session.mount('https://', TLSAdapter())
        
        response = session.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')

        print(f"Extracting information from {url}")
        
        # Extract emails
        emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', soup.text)
        print(f"Emails found: {emails}")
        
        # Enhanced regex for British phone numbers
        phone_pattern = re.compile(r'''
            (?:\+44|0044|0)\s?\d{4}\s?\d{6}      # Matches numbers like +44 1234 567890, 0044 1234 567890, 01234 567890
            |
            (?:\+44|0044|0)\s?7\d{3}\s?\d{6}     # Matches mobile numbers like +44 7123 456789, 0044 7123 456789, 07123 456789
        ''', re.VERBOSE)
        
        phone_numbers = phone_pattern.findall(soup.text)
        print(f"Phone numbers found: {phone_numbers}")
        
        # Extract company name
        company_name = extract_company_name(soup)
        print(f"Company name found: {company_name}")
        
        # If both email and phone number are absent, return None
        if not emails and not phone_numbers:
            return None
        
        return {'url': url, 'emails': emails, 'phone_numbers': phone_numbers, 'names': company_name}
    except Exception as e:
        print(f"Error extracting info from {url}: {e}")
        return None

def extract_company_name(soup):
    # Attempt to extract the company name from the title tag or meta tags
    title_tag = soup.find('title')
    if title_tag and title_tag.string:
        return title_tag.string.strip()

    # Fallback to meta tags for company name
    meta_og_title = soup.find('meta', property='og:title')
    if meta_og_title and meta_og_title.get('content'):
        return meta_og_title.get('content').strip()

    meta_title = soup.find('meta', {'name': 'title'})
    if meta_title and meta_title.get('content'):
        return meta_title.get('content').strip()

    # Fallback to a placeholder if no title or meta tags are found
    return 'Unknown'

def save_to_excel(data, filename='results.xlsx'):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['URL', 'Emails', 'Phone Numbers', 'Names'])
    
    for entry in data:
        emails = ', '.join(entry['emails'])
        phone_numbers = ', '.join(entry['phone_numbers'])
        names = entry['names']
        ws.append([entry['url'], emails, phone_numbers, names])
    
    wb.save(filename)

def main():
    query = input("Enter your search query: ")
    links = search_google(query)
    data = []

    for link in links:
        info = extract_info(link)
        if info:
            data.append(info)
    
    save_to_excel(data)

if __name__ == "__main__":
    main()
