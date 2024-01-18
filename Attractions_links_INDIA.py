from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import openpyxl
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium_stealth import stealth

# Function to initialize a WebDriver
def get_driver(url, wait_for):
    options = webdriver.ChromeOptions()
    options.add_argument("start-maximized")
    # options.add_argument("--headless")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)

    options.add_experimental_option(
        "prefs", {
            # block image loading
            "profile.managed_default_content_settings.images": 2,
            # block JavaScript
            'profile.managed_default_content_settings.javascript': 2,
        }
    )
    options.add_experimental_option("detach", True)
    options.binary_location = './/chrome-win64/chrome.exe'

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    stealth(driver,
            languages=["en-US", "en"],
            vendor="Google Inc.",
            platform="Win32",
            webgl_vendor="Intel Inc.",
            renderer="Intel Iris OpenGL Engine",
            fix_hairline=True,
            )

    driver.get(url)

    if len(wait_for) > 0:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, f".{wait_for}"))
        )

    return driver


# Output Excel file path
path = r'D:\attraction_link_india.xlsx'
wb = openpyxl.Workbook()
sh1 = wb.active
sh1.append(['tour', 'link'])  # Adding headers to the Excel file

url = 'https://www.tripadvisor.com/Attractions-g293860-Activities-oa0-India.html'

# Initialize Chrome WebDriver with options to ignore loading images and block JavaScript
# chrome_options = Options()

driver = get_driver(url, "")
driver.maximize_window()

# Set to store unique links
unique_links = set()

while True:
    try:
        # Wait for the tour titles to be present
        tour_titles = WebDriverWait(driver, 200000000000).until(
            EC.presence_of_all_elements_located((By.XPATH, '//div[@class="XfVdV o AIbhI"]'))
        )
        tour_links = driver.find_elements(By.XPATH, '//div[@class="alPVI eNNhq PgLKC tnGGX"]/a[1]')

        # Iterate through each element to get its text and save to Excel
        for tour, link in zip(tour_titles, tour_links):
            tour_name = tour.text if tour else 'N/A'
            link_href = link.get_attribute('href') if link else 'N/A'

            # Check if the link is not in the set to avoid duplicates
            if link_href not in unique_links:
                unique_links.add(link_href)
                sh1.append([tour_name, link_href])

        # Check if there is a next page
        next_page = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@class="xkSty"]'))
        )
        # Click the next page button
        next_page.click()
        # Wait for the page to load
        WebDriverWait(driver, 20).until(
            EC.staleness_of(tour_titles[0])
        )
    except TimeoutException:
        # Exit the loop if there is no next page
        page_title = driver.find_element(By.XPATH, '//div[@class="Ci"]').text
        print(page_title)
        break

# Save the workbook
wb.save(path)
# Close the WebDriver when done
driver.quit()