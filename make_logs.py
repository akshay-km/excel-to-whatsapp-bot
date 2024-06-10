import os
from datetime import datetime

class LOGGER():
    def __init__(self):
        self.file_path = None


    def create_logfile(self):
        logdir = "logs"
        if not os.path.exists(logdir):
            os.makedirs(logdir)

        current_time = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        file_name = f"log_{current_time}.txt"
        self.file_path = os.path.join(logdir, file_name)

        with open(self.file_path, 'w') as log_file:
            log_file.write(f"Log file created at {current_time} \n")


    def log(self, message: str):
        if self.file_path == None:
            self.create_logfile()

        with open(self.file_path, 'a') as log_file:
            log_file.write(message+"\n")


#GLOBAL VARIABLE to maintain same LOGGER instance across mulitple function calls
logger = None

def get_logger():
    global logger
    if logger is None:
        logger = LOGGER()
    return logger