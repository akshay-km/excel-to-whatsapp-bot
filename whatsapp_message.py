import os
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import urllib.parse
import pandas as pd
import threading
from queue import Queue
from selenium.common.exceptions import TimeoutException
from make_logs import get_logger


def open_whatsapp(driver, log_callback):
    print("Opening whatsapp for the first time...")
    log_callback("Opening whatsapp for the first time...")
    logger.log("Opening whatsapp for the first time...")
    driver.get("https://web.whatsapp.com")

    try:
        WebDriverWait(driver, 300).until(
            EC.presence_of_element_located((By.XPATH, '//canvas[@aria-label="Scan me!"]'))
        )
        print("Please scan the QR code to log in.")
        log_callback("Please scan the QR code to log in.")
        logger.log("Please scan the QR code to log in.")
        WebDriverWait(driver, 120).until_not(
            EC.presence_of_element_located((By.XPATH, '//canvas[@aria-label="Scan me!"]'))
        )
    except Exception as e:
        print(f"Log in Error : {e}")
        log_callback("Unable to login to whatsapp. Check internet connection.")
        logger.log(f"Whatsapp log in error: {e}")
        return False
    else:
        print("Successfully logged in to WHATSAPP.")
        log_callback("Successfully logged in to WHATSAPP.")
        logger.log("Successfully logged in to WHATSAPP.")
        
    print("Reloading Whatsapp chats page...")
    logger.log("Reloading Whatsapp chats page....")
    driver.get("https://web.whatsapp.com")
    try:
        print("Waiting for the chat page to load...")
        log_callback("Waiting for the chat page to load...")
        logger.log("Waiting for the chat page to load...")
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//div[@aria-label="Chat list"]'))
        ) 
    except Exception as e:
        print (f"Error loading chat age : {e}")  
        log_callback("Error loading chat page. ")
        logger.log (f"Error loading chat age : {e}")  
        return False

    else: 
        print("Chat page loaded successfully")
        log_callback("Chat page loaded successfully :)")
        logger.log("Chat page loaded successfully")
        return True


error_list = []
success_count = 0


def add_to_failed_list(name, phone_number, message, link, remarks ):
    error_list.append((phone_number, message, link, name, remarks))


# Function to check for invalid number pop-up
def check_invalid_number_popup(driver, stop_event, result_queue):
    print("CHECKING Invalid number pop up")
    logger.log("Waiting for INVALID NUMBER POP UP ")
    timeout = 30  # total timeout in seconds
    check_interval = 1  # interval in seconds to check send_button_event
    start_time = time.time()
    try:
        while True:
            if time.time() - start_time >= timeout:
                raise TimeoutException("Timed out waiting for Invalid number pop up")
            if stop_event['flag']:
                print("Send button loaded, Stopping wait for INVALID POP UP ")
                logger.log("Send button loaded, Stopping wait for INVALID POP UP ")
                return           
            try:
                WebDriverWait(driver, check_interval).until(
                    EC.presence_of_element_located((By.XPATH, '//div[contains(text(), "Phone number shared via url is invalid")]'))
                    )  
                with stop_event['lock']:
                    stop_event['flag'] = True
                print("This phone number is not registered on WhatsApp.")
                logger.log("INVALID NUMBER : This phone number is not registered on WhatsApp.")
                result_queue.put(("INVALID_POP_UP"))
                return
            except TimeoutException:
                # If we timed out, just loop again
                continue
    except Exception as e:
        print(f"Error while waiting for invalid pop up: {e}")
        logger.log(f"Error while waiting for invalid pop up: {e}")
        result_queue.put("POP_UP_TIMEOUT")
# Function to wait for the send button and click it
def wait_for_send_button(driver, stop_event, result_queue):
    print("Waiting for send button....")
    logger.log("Waiting for send button....")
    timeout = 60  # total timeout in seconds
    check_interval = 1  # interval in seconds to check pop_up_event
    start_time = time.time()

    try:
        while True:
            if time.time() - start_time >= timeout:
                raise TimeoutException("Timed out waiting for send button")
            
            if stop_event['flag']:
                print("Pop up loaded, stopping wait for SEND BUTTON")
                logger.log("Pop up loaded, stopping wait for SEND BUTTON")
                # result_queue.put(("send_button", None))
                return
            
            try:
                send_button = WebDriverWait(driver, check_interval).until(
                    EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Send"]'))
                )
                driver.execute_script("arguments[0].click();", send_button)
                print("SEND BUTTON clicked. Waiting for MESSAGE STATUS...")
                logger.log("SEND BUTTON clicked. Waiting for MESSAGE STATUS...")
                # Wait for the blue ticks to appear next to the most recent message sent
                try: 
                    blue_ticks_xpath = '//div[contains(@class, "message-out")][last()]//span[@aria-label=" Sent " or @aria-label=" Read " or @aria-label=" Delivered "]'
                    blue_ticks = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.XPATH, blue_ticks_xpath))
                    )
                    print("Message Status : SENT")
                    logger.log("Message Status : SENT")
                    result_queue.put("MESSAGE_SENT")
                except:
                    print("COULD NOT FETCH MESSAGE STATUS ")
                    logger.log("COULD NOT FETCH MESSAGE STATUS ")
                    result_queue.put("MESSAGE_NO_STATUS")

                with stop_event['lock']:
                    stop_event['flag'] = True
                # result_queue.put("")
                return
            except TimeoutException:
                # If we timed out, just loop again
                continue
    except Exception as e:
        print(f"Error while waiting for send button: {e}")
        logger.log(f"Error while waiting for send button : {e}")
        result_queue.put("SEND_BUTTON_TIMEOUT")


def open_whatsapp_and_send_message(driver, phone_number, message, name, log_callback):
    print(f"\n Name: {name} Phone: {phone_number} Message: {message}")
    log_callback(f"Name: {name} Phone: {phone_number} Message: {message}")
    logger.log(f"Name: {name} Phone: {phone_number} Message: {message}")
    # URL encode the message
    encoded_message = urllib.parse.quote(message)
    # WhatsApp Web URL
    url = f"https://web.whatsapp.com/send?phone={phone_number}&text={encoded_message}"

    driver.get(url)

    # Shared flag with thread lock 
    stop_event = {'flag': False, 'lock': threading.Lock()} 

    result_queue = Queue()

    # Start threads for each check
    pop_up_thread = threading.Thread(target=check_invalid_number_popup, args=(driver, stop_event, result_queue))
    button_thread = threading.Thread(target=wait_for_send_button, args=(driver, stop_event, result_queue))

    pop_up_thread.start()
    button_thread.start()
    print("Threads started")
    logger.log("Threads started")

    # Stop the threads
    pop_up_thread.join(timeout=30)
    button_thread.join(timeout=65)

    results_list = []

    while not result_queue.empty():
        result = result_queue.get()
        results_list.append(result)

    print(f"RESULTS_QUEUE : {results_list}")
 
    if 'MESSAGE_SENT' in results_list:
        global success_count
        success_count += 1
        log_callback(f"====== MESSAGE SENT SUCCESSFULLY ======>  {success_count}  ")
        print(f"======   MESSAGE SENT SUCCESSFULLY ======>  {success_count}")
        logger.log("====== 1 MESSAGE SENT SUCCESSFULLY ======>  {success_count} ")
        
    elif 'MESSAGE_NO_STATUS' in results_list:
        log_callback("===== MESSAGE DELIVERY STATUS : UNKNOWN. Added to Failed logs =====")
        logger.log("===== MESSAGE DELIVERY STATUS : UNKNOWN. Added to Failed logs =====")
        print("===== MESSAGE DELIVERY STATUS : UNKNOWN. Added to Failed logs =====")
        add_to_failed_list(name, phone_number, message, url, remarks="MESSAGE STATUS UNKNOWN. Please verify delivery status" )
    elif 'INVALID_POP_UP' in results_list:
        log_callback(f"===== MESSAGE FAILED {phone_number} is not registered with Whatsapp. Added to Failed logs ===== ")
        print(f"===== MESSAGE FAILED {phone_number} is not registered with Whatsapp. Added to Failed logs ===== ")
        logger.log(f"===== MESSAGE FAILED {phone_number} is not registered with Whatsapp. Added to Failed logs ===== ")
        add_to_failed_list(name, phone_number, message, url, remarks="NOT REGISTERED WITH WHATSAPP")
    elif 'SEND_BUTTON_TIMEOUT':
        print("====== FAILED TO SENT MESSAGE. Added to Failed list ==========")
        log_callback("====== FAILED TO SENT MESSAGE. Added to Failed list ==========")
        logger.log("====== FAILED TO SENT MESSAGE. Added to Failed list ==========")
        add_to_failed_list(name, phone_number, message, url, remarks="FAILED")
    else:
        print(f"RESULTS LIST : {results_list}")

def read_file(file_name: str, sheet_list: list):
    print(f"File name: {file_name}")
    print(f"Sheet list: {sheet_list}")
    logger.log(f"File name: {file_name}")
    logger.log(f"Sheet list: {sheet_list}")
    try:
        df_dict = pd.read_excel(file_name, sheet_name=sheet_list, skiprows=1, dtype={'MOB': str})
        # Reading multiple sheets returns a dict with key:value pair as name:df
        df = pd.concat(df_dict.values(), ignore_index=True)    
    except FileNotFoundError:
        print(f"Cannot find a file name {file_name}.")
        logger.log(f"Cannot find a file name {file_name}.")
    except ValueError:
        print(f"Cannot find the sheets with the given name. Try again.")
        logger.log(f"Cannot find the sheets with the given name. Try again.")
    except Exception as e:
        print(f"Error : {e}")
        logger.log(f"Error : {e}")
    else:
        required_headers = ['NAME','MOB','MESSAGE']
        missing_headers = []
        for head in required_headers:
            if head not in df.columns:
                missing_headers.append(head)
        if len(missing_headers) > 0:
            print(f"if missing headers: {missing_headers} ")
            return {"headers_missing": missing_headers}
        df.columns = [*df.columns[:-1], 'ACTION']
        # df.dropna(subset=['ACTION'], inplace=True)
        df = df[df['ACTION'] == 'SEND']
        return {"df": df}
    
def send_message(df, driver, log_callback, done_callback):
    print("Sending messages..")
    logger.log("Sending messages..")
    log_callback("Sending messages...")

    for row in df.itertuples(index=False):
        if row.ACTION == 'SEND':
            # log_callback(f"Sending message to {row.MOB}")
            try:
                open_whatsapp_and_send_message(driver, row.MOB, row.MESSAGE, row.NAME, log_callback)
            except Exception as e:
                print(f"Exception in send messages : {e}")
                logger.log(f"Error occurred in send messages : {e}")


def get_excel_file_path() -> str:
    directory = "FAILED"
    if not os.path.exists(directory):
        os.makedirs(directory)

    now = datetime.now()
    date_time_str = now.strftime("%d-%m-%Y-%H-%M-%S")
    file_name = f"failed-{date_time_str}.xlsx"
    file_path = os.path.join(directory, file_name)
    return file_path


def create_error_excel(error_list, log_callback):
    print(f"FAILED to sent {len(error_list)} messages!")
    ph = []
    msg = []
    urls = []
    names = []
    remarks = []
    for error in error_list:
        ph.append(error[0])
        msg.append(error[1])
        urls.append(error[2])
        names.append(error[3])
        remarks.append(error[4])

    df1 = pd.DataFrame({
        "NAME": names,
        "PHONE NUMBER": ph,
        "MESSAAGE": msg,
        "LINK": urls,
        "REMARKS": remarks
    })
    try:
        file_path = get_excel_file_path()
        logger.log(f"Excel file path : {file_path}")
        writer = pd.ExcelWriter(file_path, engine="xlsxwriter")

        # Convert dataframe to XlsxWriter Excel object
        df1.to_excel(writer, sheet_name='Sheet1', index=False)
        worksheet = writer.sheets['Sheet1']

        # Iterate through the DataFrame and add hyperlinks
        for row_num, link in enumerate(df1['LINK'], start=1):
            worksheet.write_url(row_num, 3, link, string="SEND")

        writer.close()
    except Exception as e:
        print(f"An error occurred : {e}")
        logger.log(f"An error occurred while creating Excel file  : {e}")
        log_callback("An error occurred while creating Failed messages excel file !")
        msg = "An error occurred while creating Failed messages excel file !"
    else:
        print("Failed messages written to failed folder")
        logger.log("Failed messages written to failed folder")
        log_callback("Failed messages written to failed folder!")
        msg = "Failed messages written to failed folder!"
    return msg


class UnsupportedOs(Exception):
    def __init__(self, message):
        self.message = message
        super().__init__(self.message)


def get_chrome_driver(log_callback):
    if os.name == "nt":
        # Chrome - Windows => Add path of chromedriver to Env PATH
        print("Getting Chromedriver for WINDOWS")
        log_callback("Getting Chromedriver for WINDOWS")
        logger.log("Getting Chromedriver for WINDOWS")
        chrome_driver = webdriver.Chrome()
        return chrome_driver
    elif os.name == "posix":
        # Chrome - Linux
        print("Getting Chromedriver for LINUX")
        log_callback("Getting Chromedriver for LINUX")
        logger.log("Getting Chromedriver for LINUX")

        chrome_options = Options()
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_driver = webdriver.Chrome(options=chrome_options)
        return chrome_driver
    else:
        print("Invalid OS; only works on  windows or linux")
        logger.log("Invalid OS; only works on  windows or linux")
        raise UnsupportedOs(f"This apps only supports WINDOWS and LINUX")


def get_log_file_path() -> str:
    directory = "RESULTS"
    if not os.path.exists(directory):
        os.makedirs(directory)

    now = datetime.now()
    date_time_str = now.strftime("%d-%m-%Y-%H-%M-%S")
    file_name = f"completed-{date_time_str}.txt"
    file_path = os.path.join(directory, file_name)
    return file_path


def log_result(excel_file_name: str, sheet_list: list, message: str, log_callback):
    file_path = get_log_file_path()
    logger.log(f"Results file path : {file_path}")
    lines = [
        f"File Name : {excel_file_name} \n",
        f"Selected Sheets : {sheet_list} \n",
        f"{message} \n"
    ]
    try:
        with open(file_path, 'w') as file:
            file.writelines(lines)
        log_callback("Results written to completed folder.")
        logger.log("Results written to completed folder.")        
    except Exception as e:
        print(f" Error occurred write logs: {e}")
        log_callback("Error occurred while writing logs to completed folder")
        logger.log("Error occurred while writing logs to completed folder")


def logout_whatsapp(driver, log_callback):
    print("Trying to log out of Whatsapp...")
    logger.log("Trying to log out of Whatsapp...")
    log_callback("Trying to log out of Whatsapp...")
    # Locate and click the menu (three dots) button
    try:
        menu_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@title="Menu"]'))
        )
        menu_button.click()

        # Locate and click the "Log out" option
        logout_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@aria-label="Log out"]'))
        )
        logout_button.click()
        # Wait for the confirmation popup and click the "Log out" button
        confirm_logout_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@class="x1c4vz4f xs83m0k xdl72j9 x1g77sc7 x78zum5 xozqiw3 x1oa3qoh x12fk4p8 x3pnbk8 xfex06f xeuugli x2lwn1j xl56j7k x1q0g3np x6s0dn4" and text()="Log out"]'))
        )
        confirm_logout_button.click()

        # Wait till login page reappears
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//canvas[@aria-label="Scan me!"]'))
        )
    except Exception as e: 
        print(f"Unable to logout : {e}")
        log_callback(f"Unable to logout : {e}")
        logger.log(f"Unable to logout : {e}")

    print("Logged out of WhatsApp Web")
    log_callback("Logged out of WhatsApp Web")
    logger.log("Logged out of WhatsApp Web !")

def start_whatsapp_messaging(file_name: str, sheet_list: list, log_callback, done_callback):

    #INITIALIZE Logger
    global logger
    logger = get_logger()

    excel_result = read_file(file_name, sheet_list)
    print(f"excel result : {excel_result}")
    if 'headers_missing' in excel_result.keys():
        print(f"The following columns are missing in the selected sheets: {excel_result['headers_missing']}")
        log_callback(f"The following columns are missing in the selected sheets: {excel_result['headers_missing']}")
        log_callback("Add these columns to all the sheets and try again. ")
        logger.log(f"COLUMNS MISSING :  {excel_result['headers_missing']}")
        done_callback("INVALID EXCEL SHEET", f"The following columns are missing in the selected sheets: {excel_result['headers_missing']}",
                      "red")
        log_callback("Click on X to close")
        return
    # elif 'df' in excel_result.keys():
    else:
        df = excel_result['df']
        if df.shape[0] == 0:
            print("The given sheet has 0 messages to be send. Make sure to add 'SEND' as value for last column in each row.")
            logger.log("The given sheet has 0 messages to be send. Make sure to add 'SEND' as value for last column in each row.")
            log_result(file_name, sheet_list, "TOTAL MESSAGES TO SEND : 0", log_callback)
            log_callback("The given sheet has 0 messages to be send. Make sure to add 'SEND' as value for last column in each row.")
            done_callback("ZERO MESSAGES TO SEND",
                           "The given sheet has 0 messages to be send. Make sure to add 'SEND' as value for last column in each row to be send.",
                          "red")
            log_callback(">>> Click X to close the application. <<<")
            return

    try:
        driver = get_chrome_driver(log_callback)
    except Exception as e:
        log_callback(f"Chrome driver Error ! ")
        logger.log(f"Chrome driver Error: {e} ")
        done_callback("CHROMEDRIVER ERROR",
                      "Make sure you have added chromedriver directory to PATH environment variable"
                      )
    else:
        opened = open_whatsapp(driver, log_callback)
        if not opened:
            print("UNABLE TO LOG IN TO WHATSAPP ! CLOSE THE PROGRAM AND RUN AGAIN.")
            log_callback(">>> UNABLE TO LOG IN TO WHATSAPP ! CLOSE THE PROGRAM AND RUN AGAIN. <<<")
            logger.log(">>> UNABLE TO LOG IN TO WHATSAPP ! CLOSE THE PROGRAM AND RUN AGAIN. <<<")
            done_callback("WHATSAPP LOGIN ERROR", "UNABLE TO LOGIN TO WHATSAPP", "red")

            print("Closing browser.")
            logger.log("Closing browser.")
            driver.quit()
            return
        
        send_message(df, driver, log_callback, done_callback)

        if success_count == df.shape[0] and len(error_list) == 0:
            print("All MESSAGES SENT SUCCESSFULLY")
            log_callback("All MESSAGES SENT SUCCESSFULLY")
            logger.log("All MESSAGES SENT SUCCESSFULLY")
            msg = "All MESSAGES SENT SUCCESSFULLY"
        else:
            msg = create_error_excel(error_list, log_callback)

        result_text = f"TOTAL MESSAGES: {df.shape[0]} | SUCCESS : {success_count} | FAILED: {len(error_list)}"
        print(result_text)
        log_callback(result_text)
        logger.log(result_text)
        
        # Logout of whatsapp
        logout_whatsapp(driver, log_callback)

        # Close the browser
        driver.quit()
        log_result(file_name, sheet_list, result_text, log_callback)
        
        print("<<< MESSAGING COMPLETED >>>")
        logger.log("<<< MESSAGING PROCESS COMPLETE. CLICK X button TO CLOSE THE APP >>>")
        log_callback("<<< MESSAGING PROCESS COMPLETE. CLICK X button TO CLOSE THE APP >>>")
        done_callback("MESSAGING COMPLETED", result_text + f"\n {msg}")

