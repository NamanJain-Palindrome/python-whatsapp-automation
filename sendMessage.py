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
