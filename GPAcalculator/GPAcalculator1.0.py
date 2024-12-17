'''     DO BEFORE RUNNING THE CODE
In the terminal, input the following commands, replacing 
"your_email" and "your_password" with your actual email and password:

Linux/macOS:
export SENECA_EMAIL="your_email"
export SENECA_PASSWORD="your_password"

Windows:
set SENECA_EMAIL=your_email
set SENECA_PASSWORD=your_password
'''
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager
import time
import os

# Function to click a button by ID with error handling
def click_button_by_id(driver, element_id, wait, button_name="button"):
    try:
        button = wait.until(EC.element_to_be_clickable((By.ID, element_id)))
        button.click()
        print(f"Clicked the {button_name} with ID: {element_id} successfully!")
    except Exception as e:
        print(f"Failed to click {button_name} with ID {element_id}: {e}")


# Function to click an element by XPath with error handling
def click_element_by_xpath(driver, xpath, wait, element_name="element"):
    try:
        element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        element.click()
        print(f"Clicked the {element_name} with XPath: {xpath} successfully!")
    except Exception as e:
        print(f"Failed to click {element_name} with XPath {xpath}: {e}")

def login_to_seneca(driver, wait, email, password):
    try:
        email_input = wait.until(EC.element_to_be_clickable((By.ID, "i0116")))
        email_input.send_keys(email)
        click_button_by_id(driver, "idSIButton9", wait, "next")

        password_input = wait.until(EC.element_to_be_clickable((By.ID, "i0118")))
        password_input.send_keys(password)
        click_button_by_id(driver, "idSIButton9", wait, "signin")

        click_button_by_id(driver, "KmsiCheckboxField", wait, "checkbox")
        click_button_by_id(driver, "idSIButton9", wait, "Yes")
        print("Login successful!")
    except Exception as e:
        print(f"Login failed: {e}")

# Function to extract and print grades for a course
def get_course_grades(driver, course_name, course_xpaths, wait):
    try:
        for xpath in course_xpaths:
            try:
                # Try to click each XPath and if successful, break out
                course_button = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                driver.execute_script("arguments[0].scrollIntoView({block: 'start'});", course_button)
                time.sleep(1)  # wait for smooth scroll
                course_button.click()
                print(f"Course '{course_name}' selected successfully!")
                break
            except Exception:
                continue  # If current XPath fails, move to the next

       

        # Get the parent item elements
        items = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//div[@ng-repeat="grade in courseGradesStudent.userGrades"]')))
        data = []

        # Iterate over each item and extract Item Name and Grade
        for item in items:
            try:
                # Extract Item Name (including link)
                item_name = item.find_element(By.XPATH, './/div[@class="name ellipsis student-grades__exception"]').text.strip()
                # Extract Grade
                 # Determine XPath for grades based on the course name (IPC or not)
                if "IPC" in course_name:  # Check if the course is IPC
                    grades_xpath = './/span[contains(@class, "grade-input-display") and contains(@class, "grade-ellipsis")]'
                else:
                    grades_xpath = './/span[@class="grade-input-display grade-ellipsis"]'
                grade = item.find_element(By.XPATH, grades_xpath).text.strip()
                data.append([item_name, grade])
            except Exception as e:
                print(f"Error processing item: {e}")

        # Print the results
        print("Grades Report:\n")
        for idx, (name, grade) in enumerate(data, start=1):
            print(f"{idx}. Item: {name}, Grade: {grade}")

        # Close the course view window
        click_element_by_xpath(driver, '//*[@id="main-content"]/div[3]/div/div[3]/button', wait, "close")


    except Exception as e:
        print(f"Failed to get grades for the course '{course_name}': {e}")

def main():
    options = Options()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get("https://learn.senecapolytechnic.ca/")
    print("Browser launched successfully!")

    wait = WebDriverWait(driver, 15)

    # Click "OK" button
    click_button_by_id(driver, "agree_button", wait, "OK")
    # Click login button
    click_button_by_id(driver, "bottom_Submit", wait, "Login")


    # Login process
    email = os.getenv("SENECA_EMAIL")
    password = os.getenv("SENECA_PASSWORD")
    login_to_seneca(driver, wait, email, password)
    
    # Wait for grades button and click
    time.sleep(3)
    click_element_by_xpath(driver, '//bb-base-navigation-button[8]//a', wait, "Grades")

    # Wait for content to load
    try:
        main_container = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-content-inner"]')))
        print("Main container located successfully.")
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", main_container)
    except Exception as e:
        print(f"Failed to locate main container: {e}")

    # Retrieve and print grades for multiple courses
    courses = [
        ('IPC', ['(//*[@id="card__733205_1"]/div/div/div[2]/div[2]/div/a)[2]']),
        ('OPS', ['//*[@id="card__731947_1"]/div/div/div[2]/div[2]/div/a/span[1]']),
        ('CPR', ['(//*[@id="card__737264_1"]/div/div/div[2]/div[2]/div/a/span[1])[2]']),
        ('APS', ['(//*[@id="card__733550_1"]/div/div/div[2]/div[2]/div/a/span[1])[2]']),
        ('COM111', ['(//*[@id="card__732569_1"]/div/div/div[2]/div[2]/div/a/span[1])[2]'])
    ]

    for course_name, course_xpaths in courses:
        print(f"\nFetching grades for {course_name}...")
        get_course_grades(driver, course_name, course_xpaths, wait)
    
    time.sleep(7)  # Let browser settle
    driver.quit()


if __name__ == "__main__":
    main()
