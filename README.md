# excel-to-whatsapp-bot
An windows app made with python , that reads excel file containing names, phone numbers , message , etc and sends message to each phone number via whatsapp web.

## REQUIREMENTS 
1.Should have Google Chrome installed.
2. Download Chromedriver version compatible to Google chrome version from (here)[https://googlechromelabs.github.io/chrome-for-testing/]. Add the path to chromedriver to the PATH environment variable.
3. The excel file must have columns with headings **NAME**, **MOB**, **MESSAGE** and **ACTION**. ACTION should be last column and for records to be send via whatsapp, the value for ACTION column must be 'SEND'.

## INSTRUCTIONS TO RUN THE PROGRAM
1. Download lastest releases and extract the folder.
2. Double-click  **whatsbot.exe** to run it.
3. Click the **Select Excel File button** and choose the excel file from the selection dialog window.
4. Names of all the sheets in the selected excel file will be displayed.
5. Select the sheet names. Multiples ones can be selected.
6. Click **Confirm Selection button**
7. Wait for browser to open, when the whatsapp website loads, scan the QR code via whatsapp.
8. Now wait till all the messages are sent.A pop up window showing number of SUCCESSFUL and FAILED messages will appear.
9. The failed entries will be written to **failed** folder in the same directory of the application. Also the results will be logged on the **completed** folder.


## LIBRARIES USED 
1. SELENIUM
2. PANDAS
3. OPENPYXL
4. XLSXWRITER
5. PYINSTALLER
