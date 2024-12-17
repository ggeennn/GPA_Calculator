from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
import logging
import os
import time

# Functions: click_button, login_to_seneca, get_courses, get_course_grades
def click_element(driver, by, identifier, wait, element_name="element"):
    try:
        element = wait.until(EC.element_to_be_clickable((by, identifier)))
        element.click()
        logging.info(f"Clicked {element_name} successfully!")
    except Exception as e:
        logging.error(f"Failed to click {element_name}: {e}")

def login_to_seneca(driver, wait, email, password):
    try:
        email_input = wait.until(EC.element_to_be_clickable((By.ID, "i0116")))
        email_input.send_keys(email)
        click_element(driver, By.ID, "idSIButton9", wait, "next")

        password_input = wait.until(EC.element_to_be_clickable((By.ID, "i0118")))
        password_input.send_keys(password)
        click_element(driver, By.ID, "idSIButton9", wait, "signin")
        click_element(driver, By.ID, "KmsiCheckboxField", wait, "checkbox")
        click_element(driver, By.ID, "idSIButton9", wait, "Yes")
        logging.info("Login successful!")
    except Exception as e:
        logging.error(f"Login failed: {e}")

def get_courses(driver, wait):
    try:
        course_elements = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//bb-base-grades-student')))
        courses = []
        for course in course_elements:
            try:
                name = course.find_element(By.XPATH, './/h3').text.strip()
                link = course.find_element(By.XPATH, './/a[contains(text(), "View all work")]').get_attribute("href") # not working
                courses.append((name, link))
            except Exception as e:
                logging.error(f"Failed to fetch course details: {e}")
                continue
        return courses
    except Exception as e:
        logging.error(f"Error fetching courses: {e}")
        return []

def get_course_grades(driver, course_name, wait):
    try:
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
                logging.error(f"Error processing item: {e}")

        # Print the results
        print("Grades Report:\n")
        for idx, (name, grade) in enumerate(data, start=1):
            print(f"{idx}. Item: {name}, Grade: {grade}")

        # Close the course view window
        click_element(driver, By.XPATH, '//*[@id="main-content"]/div[3]/div/div[3]/button', wait, "close")
    except Exception as e:
        logging.error(f"Failed to get grades for the course '{course_name}': {e}")

def main():
    # Setup WebDriver
    options = Options()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 15)
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    try:
        driver.get("https://learn.senecapolytechnic.ca/")
        logging.info("Browser launched successfully!")
  
        # Click "OK" button
        click_element(driver, By.ID, "agree_button", wait, "OK")
        # Click login button
        click_element(driver, By.ID, "bottom_Submit", wait, "Login")

        # Login
        email = os.getenv("SENECA_EMAIL")
        password = os.getenv("SENECA_PASSWORD")
        login_to_seneca(driver, wait, email, password)

        # 进入grades分页面
        time.sleep(3)
        click_element(driver, By.XPATH, '//bb-base-navigation-button[8]//a', wait, "Grades")
        main_container = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-content-inner"]')))
        logging.info("Main container located successfully.")
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", main_container)

        # Fetch and process courses
        courses = get_courses(driver, wait)
        for course_name, course_link in courses:
            logging.info(f"Fetching grades for {course_name}...")
            driver.get(course_link)
            get_course_grades(driver, course_name, wait)
    
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
