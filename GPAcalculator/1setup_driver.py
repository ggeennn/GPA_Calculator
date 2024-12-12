from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import StaleElementReferenceException
# from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import time

options = Options()
'''无法自动填充！！！
options.add_argument("--disable-blink-features=AutomationControlled")  # 防止Chrome的自动化标志
options.add_argument("--disable-save-password-bubble")  # 禁用保存密码的提示
options.add_argument("--autofill")  # 启用自动填充功能
'''
options.add_argument("--disable-blink-features=AutomationControlled")  # 禁用 Chrome 自动化标志
options.add_argument("--start-maximized")  # 启动时最大化浏览器

# Service 对象代表 ChromeDriver 进程
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
# driver.implicitly_wait(15)  #隐式等待元素加载（不建议和显式混用）
wait = WebDriverWait(driver, 15) # 显式等待元素加载

driver.get("https://learn.senecapolytechnic.ca/")
print("Browser launched successfully!")

# 定位并点击“OK”按钮
try:
    button = driver.find_element(By.ID, "agree_button") 
    button.click()  # 点击按钮
    print("Clicked the 'OK' button successfully!")
except Exception as e:
    print(f"Failed to click the button: {e}")

# 点击登录页面中的 Login 按钮
try:
    login_button = driver.find_element(By.ID, "bottom_Submit")  
    login_button.click()
    print("Clicked the 'Login' button successfully!")
except Exception as e:
    print(f"Failed to click 'Login' button: {e}")

try:
    email_input = wait.until(EC.element_to_be_clickable((By.ID, "i0116"))) 
    email_input.send_keys("ywang841@myseneca.ca")
    next_button = driver.find_element(By.ID, "idSIButton9")  # 定位 Next 按钮
    next_button.click()
    print("input email successfully!")
except Exception as e:
    print(f"Failed to : {e}")

try:
    password_input = wait.until(EC.element_to_be_clickable((By.ID, "i0118"))) 
    password_input.send_keys("Wyc406/304")
    signin_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))  # 显式等待元素加载（ID重名）
    signin_button.click()
    print("input password successfully!")
except Exception as e:
    print(f"Failed to : {e}")

try:
    # 等待并定位复选框
    checkbox = WebDriverWait(driver, 25).until(EC.element_to_be_clickable((By.ID, "KmsiCheckboxField")))
    if not checkbox.is_selected():  # 如果复选框未被选中
        checkbox.click()
        print("Checkbox selected successfully!")
    else:
        print("Checkbox was already selected!")

    # 等待并点击 "Yes" 按钮
    yes_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
    yes_button.click()
    print("Clicked 'Yes' button successfully!")
except Exception as e:
    print(f"Error during checkbox selection or 'Yes' button click: {e}")

try: 
    courses_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//bb-base-navigation-button[4]//a")))
    courses_button.click()
    print("'Courses' click successfully!")
except Exception as e:
    print(f"Failed to : {e}")


try:
    main_container = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-content-inner"]')))
    print("main_container located successfully.")

    # 检查容器是否有可滚动内容
    scroll_height = driver.execute_script("return arguments[0].scrollHeight;", main_container)
    client_height = driver.execute_script("return arguments[0].clientHeight;", main_container)
    print(f"scrollHeight: {scroll_height}, clientHeight: {client_height}")

    # 滚动页面，确保懒加载的内容加载完成
    driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", main_container)
    time.sleep(3)  # 等待内容加载
    print("scrollHeight after scrolling:", driver.execute_script("return arguments[0].scrollHeight;", main_container))

except Exception as e:
    print(f"Error during scrolling or element interaction: {e}")
    # 若发生 StaleElementReferenceException 错误，重新定位元素并尝试滚动
    if isinstance(e, selenium.common.exceptions.StaleElementReferenceException):
        main_container = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-content-inner"]')))
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", main_container)
        print("Scrolled main container after re-locating it.")


'''
try:
    main_container = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-content-inner"]')))
    # 确认容器是否有足够内容需要滚动
    print("Right Container height:", main_container.size)
    print("Right Container scrollHeight:", main_container.get_attribute("scrollHeight"))
    # 使用 ActionChains 模拟滚动
    actions = ActionChains(driver)
    actions.move_to_element(main_container).click_and_hold().move_by_offset(0, 300).release().perform()  # 向下滑动 300 像素
    print("Scrolled main container using Action Chains!")
    time.sleep(2)  # 等待内容加载
    # 验证新内容是否加载
    updated_scroll_height = int(main_container.get_attribute("scrollHeight"))
    print(f"Updated scrollHeight: {updated_scroll_height}")

except Exception as e:
    print(f"Failed to scroll main container using Action Chains: {e}")
'''


try: 
    IPC_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="course-link-_733205_1"]')))
    driver.execute_script("arguments[0].scrollIntoView();", IPC_button) 
    IPC_button.click()
    print("'IPC' select successfully!")
except Exception as e:
    print(f"Failed to : {e}")


time.sleep(3)

 #关闭浏览器
driver.quit()
