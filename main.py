from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import pandas as pd
from datetime import datetime
import re
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import os
import shutil
import time
from selenium.common.exceptions import NoSuchFrameException


def convert_date_format(date_value):
    try:
        # Split the date by '/'
        day, month, year = date_value.split('/')

        # Return the date in 'yyyy.mm.dd' format
        return f"{year}.{month}.{day}"
    except ValueError:
        return "Invalid date format"

# Define a custom expected condition
class wait_for_new_window(object):
    def __init__(self, current_handles):
        self.current_handles = current_handles

    def __call__(self, driver):
        return len(driver.window_handles) > len(self.current_handles)


# Define a custom function to check for the frame by index
def frame_by_index_is_available(index):
    try:
        driver.switch_to.frame(index)
        return True
    except NoSuchFrameException:
        return False


def wait_for_new_download_to_complete(temp_dir,initial_files):

    new_files = set(os.listdir(temp_dir))

    # Wait for a new file to appear
    while initial_files == new_files:
        time.sleep(1)
        new_files = set(os.listdir(temp_dir))

    new_file = list(new_files - initial_files)[0]

    # If the new file is a .crdownload file, wait for it to disappear
    if new_file.endswith('.crdownload') or new_file.endswith('.tmp'):
        while new_file in os.listdir(temp_dir):
            # print(new_file)
            time.sleep(1)


# Load the Excel file
scrutiny_list = "Assessment_List_Cir_2.xlsx"
xl = pd.ExcelFile(scrutiny_list)
df = xl.parse('Sheet1')

# df['Limitation Date/Compliance Date'] = pd.to_datetime(df['Limitation Date/Compliance Date'],format='%d-%m-%Y')
# df['Pending Since'] = pd.to_datetime(df['Pending Since'],format='%d-%m-%Y')
# df['Limitation Date/Compliance Date'] = df['Limitation Date/Compliance Date'].dt.strftime('%d-%m-%Y')
# df['Pending Since'] = df['Pending Since'].dt.strftime('%d-%m-%Y')

base_dir = 'D:\OneDrive\DCIT Central\Assessment'
temp_download_dir = r"C:\Users\179063\PycharmProjects\ITBA\temp download"

chromedriver_path = 'D:\Algo_Zerodha\chromedriver.exe'
chrome_options = Options()
chrome_options.add_experimental_option('prefs',{"download.default_directory":temp_download_dir,
                                       "download.prompt_for_download":False,
                                       "download.directory_upgrade":True,
                                       "plugins.always_open_pdf_externally":False
                                                })

service = Service(executable_path=chromedriver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)
username = "U179063"
password = "Aug$$2023"
driver.get('https://itba.incometax.gov.in')

try:
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "loginButtonClass"))
    )
except:
    print("Element not found or timeout exceeded")

driver.find_element(By.CLASS_NAME, "loginButtonClass").click()
driver.switch_to.window(driver.window_handles[-1])
passcode = str(input("Please enter a number: "))

# Convert the string to an integer
passcode = "1988" + passcode

driver.find_element("id", "username").send_keys(username)
driver.find_element("id", "password").send_keys(password)
driver.find_element("id", "passcode").send_keys(passcode)

# click login button
driver.find_element(By.CLASS_NAME, "pri_btn").click()

# Wait for a specific element to load after login
try:
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'pt123:pt_cb134'))
    )
except:
    print("Element not found or timeout exceeded")

time.sleep(5)

#Add logic to select role radio button here. but not necessary that much can be done manually

driver.find_element(By.ID, "pt123:pt_cb134").click() #role selection button

# Wait for a specific element to load after login
try:
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'pt123:pt_gsdvh5j_id_4'))
    )
except:
    print("Element not found or timeout exceeded")

# Get the current window handles
current_handles = driver.window_handles

driver.find_element(By.ID, "pt123:pt_gsdvh5j_id_4").click() #assessment link on home page
# time.sleep(5)

# Wait for a new window to appear , assessment link opens a new window
wait = WebDriverWait(driver, 10)
wait.until(wait_for_new_window(current_handles))
driver.switch_to.window(driver.window_handles[-1])

# Wait for a worklist link to become clickable after opening of new window
try:
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="menubar"]/div[4]/ul/li[2]/a[1]'))
    )
except:
    print("Element not found or timeout exceeded")

driver.find_element(By.XPATH, '//*[@id="menubar"]/div[4]/ul/li[2]/a[1]').click() #worklist click on new window

# Waits for PAN text box to located
try:
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'tbwPanTan'))
    )
except:
    print("Element not found or timeout exceeded")

# Once in worklist the main logic starts here
for index, row in df.iterrows():
    if not str(row['Scrap?']) == "Yes":
        continue
    dates_list = []
    # Get the values from columns B and C
    column_a_Value = str(row['PAN/TAN'])
    column_b_value = str(row['Name'])
    column_c_value = str(row['AY'])
    proceeding_type = str(row['Subject'])
    # print(column_a_Value+" "+column_c_value+" "+proceeding_type)

    # Create a directory for the column B value if it doesn't exist
    column_b_dir = os.path.join(base_dir, column_b_value)
    os.makedirs(column_b_dir, exist_ok=True)

    # Create a subdirectory for the column C value if it doesn't exist
    column_c_dir = os.path.join(column_b_dir, column_c_value)
    os.makedirs(column_c_dir, exist_ok=True)

    # Wait for the iframe by index to be available
    wait = WebDriverWait(driver, 10)
    wait.until(lambda driver: frame_by_index_is_available(1))  # Replace 0 with your desired index

    driver.find_element("id", "tbwPanTan").clear()
    driver.find_element("id", "tbwPanTan").send_keys(column_a_Value) #Enter PAN
    driver.find_element("id", "tbwAY1").clear()
    driver.find_element("id", "tbwAY1").send_keys(column_c_value[:4]) #Enter AY

    # Wait for a specific element to load after login
    try:
        element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "bwsearch"))
        )
    except:
        print("Element not found or timeout exceeded")

    driver.find_element("id", "bwsearch").click() #click search on worklist page

    # Wait for a specific element to load after login
    try:
        element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, r"gwNotification.gridDataList[1].propertyMap['lnkwsubject']"))
        )
    except:
        print("Element not found or timeout exceeded")

    time.sleep(1)

    link_x = driver.find_element("id", "gwNotification.gridDataList[1].propertyMap['lnkwsubject']")
    driver.execute_script("arguments[0].click();", link_x) #click the workitem link to open new page

    time.sleep(1)

    driver.switch_to.default_content()

    # Wait for the iframe by index to be available
    wait = WebDriverWait(driver, 10)
    wait.until(lambda driver: frame_by_index_is_available(2))  # Replace 0 with your desired index


    # xpath for the occasional adjournment sought by assessee pop up appearing immediately after clicking the worklist item link
    xpath_of_element = "/html/body/div[5]/div[3]/div/button"
    adjournment_flag = False
    try:
        # Wait for up to 5 seconds for the element to become available
        element = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, xpath_of_element))
        )
        # If the element is found within 5 seconds, click it
        element.click()
        adjournment_flag = True
    except:
        # Handle the exception if the element is not found within 5 seconds
        print("Element not found within 5 seconds.")


    time.sleep(1)

    driver.switch_to.default_content()

    # Wait for the iframe by index to be available
    wait = WebDriverWait(driver, 10)
    wait.until(lambda driver: frame_by_index_is_available(2))  # Replace 0 with your desired index


    # driver.find_element(By.XPATH, "/html/body/div[9]/div[3]/div/button[2]/span").click()

    # Wait for a specific element to load after login
    try:
        element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "bwSummaryOfActionWICssSec"))
        )
    except:
        print("Element not found or timeout exceeded")


    driver.find_element("id", "bwSummaryOfActionWICssSec").click()  # view case noting button

    driver.switch_to.default_content()
    wait = WebDriverWait(driver, 10)
    wait.until(lambda driver: frame_by_index_is_available(3))  # Replace 0 with your desired index

    driver.find_element("id", "bwExport").click()
    time.sleep(1)
    inial_files = set(os.listdir(temp_download_dir))
    driver.find_element(By.XPATH, "/html/body/div[9]/div[3]/div/button[2]/span").click()
    driver.switch_to.default_content()
    time.sleep(1)
    destination_dir = "D:\OneDrive\DCIT Central\Assessment\\" + column_b_value + "\\" + column_c_value
    # print(destination_dir)

    wait_for_new_download_to_complete(temp_download_dir,inial_files)

    # Overwrite files if they already exist in the destination directory
    files = os.listdir(temp_download_dir)
    for file in files:
        if file.endswith('.pdf'):
            src_path = os.path.join(temp_download_dir, file)
            dest_path = os.path.join(destination_dir, file)
            if os.path.exists(dest_path):
                os.remove(dest_path)
            shutil.move(src_path, dest_path)
        else:
            print(file+ " problem here")

    driver.switch_to.default_content()

    # Wait for the iframe by index to be available
    wait = WebDriverWait(driver, 10)
    wait.until(lambda driver: frame_by_index_is_available(3))  # Replace 0 with your desired index

    # Find all elements whose IDs contain the substring 'gwNotingHistory.gridDataList'
    potential_elements = driver.find_elements(By.XPATH, "//*[contains(@id, 'gwNotingHistory.gridDataList')]")

    # Filter elements using a regular expression
    pattern = re.compile(r'gwNotingHistory\.gridDataList\[\d+\]\.propertyMap\[\'lnkwViewDocument\']')
    matched_elements = [elem for elem in potential_elements if pattern.match(elem.get_attribute('id'))]


    for elem in matched_elements:
        if elem.text:

            # Navigate to the parent row, then to the first column, and finally to the input tag to get the date
            # date_input_elem = elem.find_element(By.XPATH, './ancestor::tr/td[1]/span/input')
            elem_text = elem.get_attribute('id')
            date_elem = elem.find_element(By.XPATH, './ancestor::tr/td[2]')
            date_value = date_elem.text
            # print(date_value)
            dates_list.append(date_value)

            converted_date = convert_date_format(date_value)
            destination_dir = "D:\OneDrive\DCIT Central\Assessment\\" + column_b_value + "\\" + column_c_value +"\\"+converted_date
            if not os.path.exists(destination_dir):
                os.makedirs(destination_dir)
            # print(destination_dir)
            initl_files = set(os.listdir(temp_download_dir))
            wait.until(EC.element_to_be_clickable((By.ID, elem_text)))
            driver.find_element(By.ID,elem_text).click()
            # elem.click()
            wait_for_new_download_to_complete(temp_download_dir,initl_files)
            # Overwrite files if they already exist in the destination directory
            files = os.listdir(temp_download_dir)
            for file in files:
                if file.endswith('.pdf'):
                    src_path = os.path.join(temp_download_dir, file)
                    dest_path = os.path.join(destination_dir, file)
                    if os.path.exists(dest_path):
                        os.remove(dest_path)
                    shutil.move(src_path, dest_path)
                else:
                    print("problem here: "+file)

            driver.switch_to.default_content()
            time.sleep(2)

            # Wait for the iframe by index to be available
            wait = WebDriverWait(driver, 10)
            wait.until(lambda driver: frame_by_index_is_available(4))  # Replace 0 with your desired index

            # driver.switch_to.frame(4) #its not a nested frame but two frames in a div
            time.sleep(1)

            # Wait for a specific element to load after login
            try:
                element = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "bwClose"))
                )
            except:
                print("Element not found or timeout exceeded")

            driver.find_element("id", "bwClose").click()

            driver.switch_to.default_content()
            time.sleep(1)

            # Wait for the iframe by index to be available
            wait = WebDriverWait(driver, 10)
            wait.until(lambda driver: frame_by_index_is_available(3))  # Replace 0 with your desired index

    # Find all elements whose IDs contain the substring 'gwNotingHistory.gridDataList'
    potential_elements_a = driver.find_elements(By.XPATH, "//*[contains(@id, 'gwNotingHistory.gridDataList')]")

    # Filter elements using a regular expression
    # pattern_a = re.compile(r'gwNotingHistory\.gridDataList\[\d+\]\.propertyMap\[\'lnkViewAttachment\']')
    # matched_elements_a = [elem_a for elem_a in potential_elements_a if pattern.match(elem_a.get_attribute('id'))]

    # Filter elements using a regular expression
    pattern_a = re.compile(r'gwNotingHistory\.gridDataList\[\d+\]\.propertyMap\[\'lnkViewAttachment\']')
    matched_elements_ids = [elem_a.get_attribute('id') for elem_a in potential_elements_a if
                            pattern_a.match(elem_a.get_attribute('id'))]

    # Now, matched_elements contains the elements with IDs that match the pattern
    for elem_id in matched_elements_ids:
        driver.switch_to.default_content()
        # Wait for the iframe by index to be available
        wait = WebDriverWait(driver, 10)
        wait.until(lambda driver: frame_by_index_is_available(3))  # Replace 0 with your desired index

        elem_aa = driver.find_element(By.ID, elem_id)
        if elem_aa.text:

            # Navigate to the parent row, then to the first column, and finally to the input tag to get the date
            # date_input_elem = elem.find_element(By.XPATH, './ancestor::tr/td[1]/span/input')
            elem_text = elem_aa.get_attribute('id')
            date_elem = elem_aa.find_element(By.XPATH, './ancestor::tr/td[2]')
            date_value = date_elem.text
            # print(date_value)
            dates_list.append(date_value)

            converted_date = convert_date_format(date_value)
            destination_dir = "D:\OneDrive\DCIT Central\Assessment\\" + column_b_value + "\\" + column_c_value + "\\" + converted_date

            if not os.path.exists(destination_dir):
                os.makedirs(destination_dir)

            # the blow code replaces driver.find_element(By.ID, elem_text).click() which mostly earlier
            elementX = driver.find_element(By.ID, elem_text)
            driver.execute_script("arguments[0].click();", elementX)


            driver.switch_to.default_content()
            # Wait for the iframe by index to be available
            wait = WebDriverWait(driver, 10)
            wait.until(lambda driver: frame_by_index_is_available(4))  # Replace 0 with your desired index

            # Find all elements whose IDs contain the substring 'gwNotingHistory.gridDataList'
            potential_elements = driver.find_elements(By.XPATH,
                                                      "//*[contains(@id, 'attachmentsGrid.gridDataList')]")

            # Filter elements using a regular expression
            pattern = re.compile(r'attachmentsGrid\.gridDataList\[\d+\]\.propertyMap\[\'lnkFileName\']')
            matched_elements2 = [elem for elem in potential_elements if pattern.match(elem.get_attribute('id'))]

            # Now, matched_elements contains the elements with IDs that match the pattern
            for elem2 in matched_elements2:
                if elem2.text:
                    initl_files = set(os.listdir(temp_download_dir))
                    elem_text2 = elem2.get_attribute('id')

                    # the below code replaces earlier code driver.find_element(By.ID, elem_text2).click() the present code uses java script
                    elementY = driver.find_element(By.ID, elem_text2)
                    driver.execute_script("arguments[0].click();", elementY)

                    # elem.click()
                    wait_for_new_download_to_complete(temp_download_dir, initl_files)

                    # Overwrite files if they already exist in the destination directory
                    files = os.listdir(temp_download_dir)
                    for file in files:
                        if file.endswith('.pdf') or file.endswith('.zip') or file.endswith('.gz'):
                            src_path = os.path.join(temp_download_dir, file)
                            dest_path = os.path.join(destination_dir, file)
                            if os.path.exists(dest_path):
                                os.remove(dest_path)
                            shutil.move(src_path, dest_path)
                        else:
                            print("problem here: " + file)

            driver.switch_to.default_content()
            # Wait for the elements with the class name "tabRight" to be present
            wait = WebDriverWait(driver, 10)  # Wait for up to 10 seconds
            elements = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "tabRight")))
            elements[-1].click()


            # this clicks the worklist tab for next iteration
            elements = driver.find_elements(By.CLASS_NAME, "tabsLi")
            elements[3].click()

            driver.switch_to.default_content()
            time.sleep(2)


    driver.switch_to.default_content()
    # time.sleep(5)

    # Wait for the elements with the class name "tabRight" to be present
    wait = WebDriverWait(driver, 10)  # Wait for up to 10 seconds
    elements = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "tabRight")))

    # Close all but the first two tabs
    # this closes all but first two tabs
    for element in elements[2:]:
        # Wait for the element to be clickable
        wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "tabRight")))
        element.click()

    #this clicks the worklist tab for next iteration
    elements = driver.find_elements(By.CLASS_NAME,"tabsLi")
    elements[-1].click()

    driver.switch_to.default_content()

    # Save the modified DataFrame back to the Excel file
    datetime_list = [datetime.strptime(date, "%d/%m/%Y") for date in dates_list]
    # latest_date = max(datetime_list)
    latest_date = max(datetime_list) if datetime_list else datetime.strptime("01/01/2000", "%d/%m/%Y")
    print("Latest date here: "+str(latest_date))
    df.at[index, 'Last activity date'] = latest_date
    if adjournment_flag:
        df.at[index, 'Script comment'] = "Ajournment sought"
    df['Last activity date'] = pd.to_datetime(df['Last activity date']).dt.strftime('%m/%d/%Y')
    df.to_excel(scrutiny_list, index=False)


# Close the browser
driver.quit()
