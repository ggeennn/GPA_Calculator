import logging
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from collections import defaultdict
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import os
import time

# Functions: click_button, login_to_seneca, get_courses, get_course_grades
def click_element(driver, by, identifier, wait, element_name="element"):
    try:
        element = wait.until(EC.element_to_be_clickable((by, identifier)))
        if element:
            # 获取元素的边界框
            rect = driver.execute_script("return arguments[0].getBoundingClientRect();", element)
            # 检查元素是否在视口内
            if rect['top'] < 0 or rect['bottom'] > driver.execute_script("return window.innerHeight;"):
                # 如果元素不在视口内，滚动到元素
                time.sleep(1) # 也可以写在scroll命令后面？？？
                driver.execute_script("arguments[0].scrollIntoView(true);", element)
                # time.sleep(1)

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

def get_course_grades(driver, course_name, wait, workbook, output_path):
    try:
        logging.info(f"Getting grades for course: {course_name}")
        # Get the parent item elements
        items = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//div[@ng-repeat="grade in courseGradesStudent.userGrades"]')))
        data = []

        # Iterate over each item and extract Item Name and Grade
        for item in items:
            try:
                # Extract Item Name
                name_xpath= './/div[@class="name ellipsis student-grades__exception"]'
                item_name = item.find_element(By.XPATH, name_xpath).text.strip()
                # Extract Grade
                grades_xpath = './/div[@class="grade-color"]'
                raw_grade = item.find_element(By.XPATH, grades_xpath).text.strip()
                score, total = map(float, raw_grade.split('/'))
                grade = score / total
                data.append([item_name, grade])
            except Exception as e:
                logging.error(f"Error processing item: {e}")

        # Print the results
        print(f"\n\n{course_name} Grades Report:")
        for idx, (name, grade) in enumerate(data, start=1):
            print(f"{idx}. Item: {name}, Grade: {grade}")
        print("\n\n")

        save_to_excel(workbook, data, course_name)

        workbook.save(output_path)
        logging.info(f"All grades saved to {output_path} successfully!")

        # Close the course view window
        click_element(driver, By.XPATH, '//*[@id="main-content"]/div[3]/div/div[3]/button', wait, "close")
    except Exception as e:
        logging.error(f"Failed to get grades for the course '{course_name}': {e}")

def normalize_name(name):
    name = name.lower()
    name = name.replace("_fall2024", "").replace("-fall2024", "").replace("fall2024", "")
    return ''.join(filter(str.isalnum, name))


def get_top_left_cell(worksheet, row, column):
    """获取合并单元格的左上角单元格"""
    cell = worksheet.cell(row=row, column=column)
    for merged_range in worksheet.merged_cells.ranges:
        if cell.coordinate in merged_range:
            min_col, min_row, max_col, max_row = merged_range.bounds
            return worksheet.cell(row=min_row, column=min_col)
    return cell

def safe_write_cell(worksheet, row, column, value):
    """安全写入单元格，避免合并单元格写入问题"""
    try:
        cell = get_top_left_cell(worksheet, row, column)
        cell.value = value
        logging.debug(f"Written value '{value}' at row {row}, column {column}")
    except ValueError as e:
        logging.error(f"Failed to write value '{value}' at row {row}, column {column}: {e}")

def save_to_excel(workbook, data, course_name):
    """保存成绩到 Excel"""
    course_mapping = {
        "OPS": {
            "columns": {
                "lab": "O",
                "quiz": "J",
                "midterm": "C",
                "final": "C"
            },
            "rows": {
                "lab_start_row": 63,
                "quiz_start_row": 63,
                "midterm_row": 66,
                "final_row": 67
            }
        },
        "CPR": {
            "columns": {
                "activity": "J",
                "quiz": "O",
                "ict": "D",
                "final": "D"
            },
            "rows": {
                "activity_start_row": 27,
                "quiz_start_row": 27,
                "ict_row": 30,
                "final_row": 31
            }
        },
        "APS": {
            "columns": {
                "test": "I",
                "ws": "O",
                "Vretta": "C",
                "BestPresentation": "C"
            },
            "rows": {
                "test_start_row": 14,
                "ws_start_row": 14,
                "Vretta_row": 20,
                "BestPresentation_row": 19
            }
        },
        "IPC": {
            "columns": {
                "ws": "J",
                "q": "O",
                "midterm": "C",
                "final": "C",
                "ms": "C"
            },
            "rows": {
                "ws_start_row": 52,
                "q_start_row": 52,
                "midterm_row": 58,
                "final_row": 59,
                "ms_start_row": 54
            }
        },
        "COM111": {
            "columns": {
                "Summary": "C",
                "Analyzing": "C",
                "Personal": "D",
                "Transfer Assignment 1": "C",
                "Transfer Assignment 2": "C",                
                "week2": "C",
                "Persuasive": "D",
                "Final": "C"
            },
            "rows": {
                "Summary_row": 41,
                "Analyzing_row": 42,
                "Personal_row": 42,
                "Transfer Assignment 1_row": 43,
                "Transfer Assignment 2_row": 44,                
                "week2_row": 45,
                "Persuasive_row": 45,
                "Final_row": 46
            }
        }
    }

    course_config = course_mapping.get(course_name, {})
    columns = course_config.get("columns", {})
    rows = course_config.get("rows", {})

    ws_averages = defaultdict(list)  # 处理 IPC 的 WS 平均分

    for item_name, grade in data:
        logging.debug(f"Processing item: {item_name}, grade: {grade}")
        matched = False
        normalized_item = normalize_name(item_name)
        logging.debug(f"normalnizing item: {normalized_item}")


        for key, column_letter in columns.items():
            normalized_key = normalize_name(key)
            if normalized_key in normalized_item:
                matched = True
                item_number = 0
                if f'{key}_start_row' in rows:
                    item_number = int(''.join(filter(str.isdigit, normalized_item)) or '0')
                if course_name == "IPC" and key == "ms":
                    if "program" in normalized_item:
                        row, column = 56, column_index_from_string("C")
                        logging.debug(f"Matched 'MS3 - Program' for {item_name}, setting row {row}, column {column}")
                    elif "video" in normalized_item:
                        row, column = 56, column_index_from_string("D")
                        logging.debug(f"Matched 'MS3 - Video' for {item_name}, setting row {row}, column {column}")
                    else:
                        row = rows.get("ms_start_row", 1) + item_number - 1
                        column = column_index_from_string(column_letter)
                        logging.debug(f"Matched 'MS' for {item_name}, setting row {row}, column {column}")
               
                elif item_number == 0:
                    row = rows.get(f"{key}_row", 1)
                    logging.debug(f"Matched key '{key}' for {item_name}, setting row {row}")
                    column = column_index_from_string(column_letter)
                else:
                    row_key = f"{key}_start_row"
                    start_row = rows.get(row_key, 1)
                    row = start_row + item_number - 1
                    logging.debug(f"Matched key '{key}' for {item_name}, setting start row {start_row} and row {row}")
                    column = column_index_from_string(column_letter)

                
                safe_write_cell(workbook["TOTAL"], row, column, grade)
                break

        if not matched:
            logging.warning(f"No matching key found for item: {item_name}")

    # 处理 IPC 的 WS 平均分
    if course_name == "IPC" and ws_averages:
        row_key = "ws_start_row"
        start_row = rows.get(row_key, 1)
        for ws_number, grades in ws_averages.items():
            average_grade = sum(grades) / len(grades)
            row = start_row + ws_number - 1
            column = column_index_from_string(columns.get("ws", "A"))
            safe_write_cell(workbook["TOTAL"], row, column, average_grade)

def main():
    # Setup WebDriver
    options = Options()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 20)
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    template_path = "SEMESTER 1 MARKINGS.xlsx"
    output_path = "Updated_Markings.xlsx"
    workbook = load_workbook(template_path)

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
        time.sleep(3) #只能写在点击命令前面？？？
        click_element(driver, By.XPATH, '//bb-base-navigation-button[8]//a', wait, "Grades")
        
        # Wait for content to load
        try:
            time.sleep(1)
            main_container = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content-inner"]')))
            logging.info("Main container located successfully.")
            scroll_height = driver.execute_script("return arguments[0].scrollHeight;", main_container)
            client_height = driver.execute_script("return arguments[0].clientHeight;", main_container)
            print(f"Scroll Height: {scroll_height}, Client Height: {client_height}")
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", main_container)
            # time.sleep(1) 也可以写在scroll命令前面？？？
        except Exception as e:
            logging.error(f"Failed to locate main container: {e}")

        # Fetch and process courses
        # courses = get_courses(driver, wait) 待优化
        courses = [
        ('IPC', '(//*[@id="card__733205_1"]/div/div/div[2]/div[2]/div/a)[2]'),
        ('OPS', '//*[@id="card__731947_1"]/div/div/div[2]/div[2]/div/a/span[1]'),
        ('CPR', '(//*[@id="card__737264_1"]/div/div/div[2]/div[2]/div/a/span[1])[2]'),
        ('APS', '(//*[@id="card__733550_1"]/div/div/div[2]/div[2]/div/a/span[1])[2]'),
        ('COM111', '(//*[@id="card__732569_1"]/div/div/div[2]/div[2]/div/a/span[1])[2]')
        ]
        for course_name, course_xpath in courses:
            logging.info(f"Fetching grades for {course_name}...")
            try:
                click_element(driver, By.XPATH, course_xpath, wait, course_name)
                get_course_grades(driver, course_name, wait, workbook, output_path)
            except Exception as e:
                logging.error(f"Select {course_name} failed: {e}")
            
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
