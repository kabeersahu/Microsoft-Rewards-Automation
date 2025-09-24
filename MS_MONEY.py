import selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook
import time
import random
import os
import psutil

# Function to check for and terminate lingering Chrome processes
def terminate_lingering_chrome_processes():
    for process in psutil.process_iter():
        if process.name() in ("chrome.exe", "chromedriver.exe") and process.username() == os.getlogin():
            try:
                process.terminate()
            except psutil.NoSuchProcess:
                pass  # Ignore if process no longer exists

# Define the list of Chrome profile directories (replace with your actual paths)
profiles = [



    r"C:\Users\ethic\AppData\Local\Google\Chrome\User Data\Profile 4",
    r"C:\Users\ethic\AppData\Local\Google\Chrome\User Data\Default",
r"C:\Users\ethic\AppData\Local\Google\Chrome\User Data\Profile 2",
r"C:\Users\ethic\AppData\Local\Google\Chrome\User Data\Profile 3",
r"C:\Users\ethic\AppData\Local\Google\Chrome\User Data\Profile 21",
r"C:\Users\ethic\AppData\Local\Google\Chrome\User Data\profile 22",
r"C:\Users\ethic\AppData\Local\Google\Chrome\User Data\profile 23",
r"C:\Users\ethic\AppData\Local\Google\Chrome\User Data\profile 24",
r"C:\Users\ethic\AppData\Local\Google\Chrome\User Data\profile 25",
r"C:\Users\ethic\AppData\Local\Google\Chrome\User Data\profile 26",
r"C:\Users\ethic\AppData\Local\Google\Chrome\User Data\profile 27",
r"C:\Users\ethic\AppData\Local\Google\Chrome\User Data\profile 28",
r"C:\Users\ethic\AppData\Local\Google\Chrome\User Data\profile 29",
r"C:\Users\ethic\AppData\Local\Google\Chrome\User Data\profile 30",
r"C:\Users\ethic\AppData\Local\Google\Chrome\User Data\Profile 32",
r"C:\Users\ethic\AppData\Local\Google\Chrome\User Data\Profile 33",








    # Add more profile paths here
]

# Path to ChromeDriver executable (replace with your actual path)
chromedriver_path = r"C:\Users\ethic\Downloads\Compressed\chromedriver-win64_2\chromedriver-win64\chromedriver.exe"

# Number of searches to perform per profile
num_searches = 21

# Load the Excel workbook containing the URLs (replace with your actual filename)
try:
    workbook = load_workbook(filename="C:\\Users\\ethic\\Music\\Book1.xlsx")
    sheet = workbook.active
except FileNotFoundError:
    print("Error: Excel file not found. Please check the path and filename.")
    exit()

# Extract the URLs from the Excel sheet (handling empty cells)
urls = [cell.value for cell in sheet["A"] if cell.value is not None]

# Shuffle the URLs to randomize the order
random.shuffle(urls)

def main():
    try:
        # Check for and terminate lingering Chrome processes
        terminate_lingering_chrome_processes()

        for profile_path in profiles:
            options = webdriver.ChromeOptions()
            options.add_argument(f"user-data-dir={profile_path}")
            options.add_argument("--start-maximized")  # Maximize window on startup

            service = Service(chromedriver_path)
            driver = webdriver.Chrome(service=service, options=options)

            for _ in range(num_searches):
                # Choose a random URL from the list
                url = random.choice(urls)
                driver.get(url)

                # Perform actions on the website (replace with your specific actions)
                try:
                    search_button = driver.find_element(By.ID, "search-button")
                    search_button.click()
                except:
                    # Handle cases where search_button is not found
                    pass

                # Wait for the search results to load
                time.sleep(5)

                try:
                    rewards_button = driver.find_element(By.XPATH, "//button[text()='Microsoft Rewards']")
                    rewards_button.click()
                except:
                    # Handle cases where rewards_button is not found
                    pass

                # Wait for some time before the next action
                time.sleep(5)

            driver.quit()

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
