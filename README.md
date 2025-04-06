# WhatsApp (Web) Automation - Python

## About the Project
This project is useful for sending **bulk messages** via WhatsApp Web from your default browser, picks the phone numbers from Microsoft Excel file and send only text message to a particular phone number.

## Built using
<img src="logos/official_python-logo-master-v3-TM-flattened.png" alt="Python Logo" width="200">

# Setting Up

## Install Python
1. Go to the official Python website: https://www.python.org/downloads/
2. Download the latest version for your operating system (Windows, macOS, or Linux).
3. Run the installer and ensure you check "Add Python to PATH" before proceeding.
4. Click Install Now and wait for the installation to complete.
5. Verify the installation by opening a terminal or command prompt and running:
   
   ```
   python --version
   ```
   OR
   ```
   python3 --version
   ```

## Install Text / Code editor
This would be the place where you would be able to modify the source code. There are many text editors available like Microsoft Visual Studio, Sublime Text, etc. We will be using Sublime Text.

1. Visit the official Sublime Text website: https://www.sublimetext.com/download
2. Select the appropriate version for your operating system.
3. Click the Download button to start the process.
4. Open the downloaded file.
5. Follow the on-screen installation instructions.
6. Choose the installation location (or leave it as default).
7. Complete the installation process by clicking Finish/Done.


## Install Python Packages
1. Open a terminal or a command prompt.
2. Install required Python packages:
   ```
   pip install pandas openpyxl pywhatkit
   ```


## Prepare your Excel file
- Ensure your Excel file *(eg. customers.xlsx)* is formatted properly with these columns:
- Make sure your *Phone Number* is in correct format. It should be of 10 digits. If there is a country code or not a country code, it does not matter, becasue the code is designed such a way that even if there is no country code, it will add country code.


## Save the Script (Source Code)
Make sure to save the the file in the same folder where your Excel file and all Python Packages are downloaded.

1. Open a text/code editor, in our case Sublime Text.
2. Click on Files (Menu Bar)
3. Select New File
4. A untitle tab would be opened.
5. Copy and paste the below code.

```py
import pandas as pd  # For handling Excel files
import pywhatkit as kit  # For sending WhatsApp messages
import time  # To add delays and prevent spamming
 
# Define the file path for the Excel sheet
file_path = "customers.xlsx"  # Update this with the actual path to your file
 
# Load the Excel file into a Pandas DataFrame
df = pd.read_excel(file_path)
 
# Check if the 'Status' column exists; if not, create it
if 'Status' not in df.columns:
    df['Status'] = ""  # Initialize with empty values
 
# Iterate through each row of the DataFrame
for index, row in df.iterrows():
    # Extract customer details from each row
    name = row["First Name"]
    phone = str(row["Phone Number"])  # Ensure phone number is treated as a string
 
    # Ensure phone number includes the correct country code
    if not phone.startswith("+"):
        phone = "+91" + phone  # Change "+91" to your country code if needed
 
    # Define the message to be sent on WhatsApp
    message = f"Good morning, {name}! ☀️\n\nAnother day to shine, crush deadlines, and pretend our coffee is working. Let’s make it a great one! ☕💪\n\nWe are excited to invite you to our program meeting. Let us know if you have any questions!\n\nBest regards,\nNaman Jain"
 
    try:
        # Send the WhatsApp message instantly
        kit.sendwhatmsg_instantly(phone, message, wait_time=20, tab_close=True)
        print(f"Message sent to {name} ({phone})")
 
        # Update the 'Status' column in the DataFrame with "Sent"
        df.at[index, 'Status'] = "Sent"
    except Exception as e:
        # If an error occurs, print the error message
        print(f"Failed to send message to {name} ({phone}): {e}")
 
        # Update the 'Status' column in the DataFrame with the failure reason
        df.at[index, 'Status'] = f"Failed - {e}"
 
    # Add a delay of 5 seconds to avoid sending messages too fast
    time.sleep(5)
 
# Save the updated Excel file with the status column updated
df.to_excel(file_path, index=False)
print("All messages processed, and status updated in the Excel sheet!")
```

6. Go to File (Menu)
7. Select Save As...
8. Save it as *sendMessage.py*

All set-ups are done, now its time to run the code/script


# How This Script Works
1. Loads an Excel file that contains customer details.
2. Checks and adds a "Status" column if it’s missing.
3. Formats phone numbers correctly by adding a country code if missing.
4. Sends a personalized WhatsApp message to each customer.
5. Logs success or failure in the "Status" column of the Excel file.
6. Saves the updated Excel file with the message delivery status.



# Important Notes
- WhatsApp Web must be logged in on your default browser.
- The script will open WhatsApp Web and send messages automatically.
- If any message fails, it will be recorded in the Excel sheet under the "Status" column.
- The updated Excel sheet will reflect "Sent" or "Failed - (error message)".
- Make sure you have a strong internet connection
- In case your laptop is slow, try to close unnecessary tabs and applications.



# Run the script
1. Open Command Promt (Windows) / Terminal (Mac/Linux)
2. Navigate to the folder where the script/code file is saved:
   ```
   cd C:\Users\Test\Documents\whatsaapp_automation
   ```
3. Run your script/code:
   ```
   python sendMessage.py
   ```
