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
from make_logs import get_logger


def open_whatsapp(driver, log_callback):
    print("Opening whatsapp for the first time...")
    log_callback("Opening whatsapp for the first time...")
    logger.log("Opening whatsapp for the first time...")
    driver.get("https://web.whatsapp.com")

    print("Please scan the QR code to log in.")
    log_callback("Please scan the QR code to log in.")
    logger.log("Please scan the QR code to log in.")
    print("Waiting for the page to load...")
    log_callback("Waiting for the page to load...")
    logger.log("Waiting for the page to load...")

    time.sleep(40)  # Delay

    driver.get("https://web.whatsapp.com")
    print("Waiting for the chat page to load...")
    log_callback("Waiting for the chat page to load...")
    logger.log("Waiting for the chat page to load...")
    time.sleep(30)


error_list = []
success_count = 0


def open_whatsapp_and_send_message(driver, phone_number, message, name, log_callback):
    print(f"\n Name: {name} Phone: {phone_number} Message: {message}")
    log_callback(f"Name: {name} Phone: {phone_number} Message: {message}")
    logger.log(f"Name: {name} Phone: {phone_number} Message: {message}")
    # URL encode the message
    encoded_message = urllib.parse.quote(message)
    # WhatsApp Web URL
    url = f"https://web.whatsapp.com/send?phone={phone_number}&text={encoded_message}"

    driver.get(url)

    try:
        # Wait until the send button is present using WebDriverWait
        send_button = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//button[@aria-label="Send"]'))
        )
        driver.execute_script("arguments[0].click();", send_button)
        global success_count
        success_count += 1
        print(f"Messages sent successfully : {success_count} ")
        log_callback(f"Messages sent successfully : {success_count} ")
        logger.log(f"Messages sent successfully : {success_count} ")
    except Exception as e:
        print(f"An error occurred: {e}")
        logger.log(f"An error occurred: {e}")
        print("Failed to send message !")
        log_callback("Failed to send message! Adding entry to failed message list")
        logger.log(f"Failed to send message! Adding entry to failed message list : {phone_number} | {message} | {url}")
        error_list.append((phone_number, message, url, name))
    finally:
        # Wait for a few seconds before proceeding
        time.sleep(5)


def read_file(file_name: str, sheet_list: list, log_callback):
    print(f"File name: {file_name}")
    print(f"Sheet list: {sheet_list}")
    logger.log(f"File name: {file_name}")
    logger.log(f"Sheet list: {sheet_list}")
    try:
        df_dict = pd.read_excel(file_name, sheet_name=sheet_list, skiprows=1, dtype={4: str})
        # Reading multiple sheets returns a dict with key:value pair as name:df
        df = pd.concat(df_dict.values(), ignore_index=True)
        df.columns = [*df.columns[:-1], 'ACTION']
        df.dropna(subset=['ACTION'], inplace=True)
        return df
    except FileNotFoundError:
        print(f"Cannot find a file name {file_name}.")
        logger.log(f"Cannot find a file name {file_name}.")
    except ValueError:
        print(f"Cannot find the sheets with the given name. Try again.")
        logger.log(f"Cannot find the sheets with the given name. Try again.")
    except Exception as e:
        print(f"Error : {e}")
        logger.log(f"Error : {e}")


def send_message(df, driver, log_callback):
    print("Sending messages..")
    logger.log("Sending messages..")
    log_callback("Sending messages...")
    for row in df.itertuples(index=False):
        if row.ACTION == 'SEND':
            # log_callback(f"Sending message to {row.MOB}")
            open_whatsapp_and_send_message(driver, row.MOB, row.MESSAGE, row.NAME, log_callback)


def get_excel_file_path() -> str:
    directory = "failed"
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
    for error in error_list:
        ph.append(error[0])
        msg.append(error[1])
        urls.append(error[2])
        names.append(error[3])

    df1 = pd.DataFrame({
        "Name": names,
        "Phone_Number": ph,
        "Message": msg,
        "Link": urls,
    })
    try:
        file_path = get_excel_file_path()
        logger.log(f"Excel file path : {file_path}")
        writer = pd.ExcelWriter(file_path, engine="xlsxwriter")

        # Convert dataframe to XlsxWriter Excel object
        df1.to_excel(writer, sheet_name='Sheet1', index=False)
        worksheet = writer.sheets['Sheet1']

        # Iterate through the DataFrame and add hyperlinks
        for row_num, link in enumerate(df1['Link'], start=1):
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
    directory = "completed"
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


def start_whatsapp_messaging(file_name: str, sheet_list: list, log_callback, done_callback):

    #INITIALIZE Logger
    global logger
    logger = get_logger()

    df = read_file(file_name, sheet_list, log_callback)
    try:
        driver = get_chrome_driver(log_callback)
    except Exception as e:
        log_callback(f"Chrome driver Error ! ")
        logger.log(f"Chrome driver Error: {e} ")
        done_callback("CHROMEDRIVER ERROR",
                      "Make sure you have added chromedriver directory to PATH environment variable"
                      )
    else:
        open_whatsapp(driver, log_callback)
        send_message(df, driver, log_callback)

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
        # Close the browser
        driver.quit()

        log_result(file_name, sheet_list, result_text, log_callback)
        logger.log("<<< MESSAGING COMPLETED >>>")
        done_callback("MESSAGING COMPLETED", result_text + f"\n {msg}")

