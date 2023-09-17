import pandas as pd 
import openpyxl
import os
import re
import win32com.client as win32
from datetime import datetime, date


def get_latest_file(folder_path, file_pattern, file_title, num_delimiters, file_type):
    # Get a list of all the files in the folder
    file_list = os.listdir(folder_path)

    # Filter file that matches the file pattern 
    matching_files = [file_name for file_name in file_list if not file_name.startswith(".")
                        and file_name.startswith(file_pattern.split("*")[0])]

    if not matching_files:
        raise ValueError(f"No files found in the folder matching the pattern {file_pattern}")

    file_numbers = []
    for file_name in matching_files:
        try:
            parts = re.split(r"[_\s]", file_name)
            file_number = int(parts[num_delimiters].split(".")[0])
            file_numbers.append(file_number)
        except (ValueError, IndexError):
            pass
    
    if not file_numbers:
        raise ValueError(f"No valid files found in the folder matching the pattern {file_pattern}")
    
    latest_number = max(file_numbers)

    if file_type == "csv":
        latest_file = f"{file_pattern.split('*')[0]}{latest_number}.csv"
    elif file_type == "xlsx":
        latest_file = f"{file_pattern.split('*')[0]}{latest_number}.xlsx"
    
    print(f"The latest {file_title} file loaded is {latest_file}.")
    
    return os.path.join(folder_path, latest_file), str(latest_number)

def count_delimiters(input_string):
    count = 0
    for char in input_string:
        if char == "_" or char == " ":
            count += 1
    return count

def check_file_type(filename):
    if filename.lower().endswith(".csv"):
        return "csv"
    elif filename.lower().endswith(".xlsx"):
        return "xlsx"
    else:
        return "Unknown"

def categorize_status(aging_days):
    if aging_days == "NA":
        return "Completed"
    if aging_days >= 1:
        return "Overdue"
    elif -6 <= aging_days <= 0:
        return "Due in < 7 days"
    else:
        return "Not Overdue"

def calculate_aging_days(row):
    if row["Payment Made"] == "No":
        return (today - row["Due date"]).days
    else:
        return "NA"


def send_email(to, cc, subject, body, attachment= None):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0) 
    mail.To = to
    mail.CC = cc
    mail.Subject = subject
    mail.HTMLBody = body
    if attachment:
        mail.Attachments.Add(attachment)
    mail.Save() # to save the email draft in the outlook without sending
    # mail.Send() # uncomment this to send out the email immediately 
    
def generate_html_table(dataframe):
    columns_to_display =["Item Bought", "Cost per unit", "Qty", "Total Cost", "Due date", "Status"]
    selected_data = dataframe[columns_to_display]
    html_table = selected_data.to_html(index=False)
    return html_table

def validate_date(date):
    try:
        datetime.strptime(date, "%d %b %Y")
        return True
    except ValueError:
        return False

# To ask for user input for deadline
while True:
    due_date = input("Please specify the vendor response deadline: ")
    if validate_date(due_date):
        break
    else:
        print("Invalid date format. Please enter the correct format '10 Jan 2022'.")

# Update Status file info
status_folder_path = r"C:\Users\Auklet\Dropbox\Cloud Projects\python email"
status_file_pattern = "Update_Status_*.csv"
status_delimiters = count_delimiters(status_file_pattern)
status_file_type = check_file_type(status_file_pattern)

# Contact List file info
contactlist_folder_path = r"C:\Users\Auklet\Dropbox\Cloud Projects\python email"
contactlist_file_pattern = "Department contact list *.xlsx"
contactlist_delimiters = count_delimiters(contactlist_file_pattern)
contactlist_file_type = check_file_type(contactlist_file_pattern)

# Search for the latest file for Upstate Status and Contact List
status_file_path, status_latest_date = get_latest_file(status_folder_path, status_file_pattern, "Update Status", status_delimiters, status_file_type)
contactlist_file_path, contactlist_latest_date = get_latest_file(contactlist_folder_path, contactlist_file_pattern, "Contact List", contactlist_delimiters, contactlist_file_type)

# Read file into pandas dataframe
df_status = pd.read_csv(status_file_path)
df_contact = pd.read_excel(contactlist_file_path)

# Choosing the required column
df_status = df_status[["S/N", "Vendor", "Item Bought", "Cost per unit", "Qty", "Due date", "Payment Made"]]
df_contact = df_contact[["Vendor", "Name", "Email address"]]

df_status["Total Cost"] = df_status["Cost per unit"] * df_status["Qty"]
# Convert the due date column to datetime objects
df_status["Due date"] = pd.to_datetime(df_status["Due date"], format="%d/%m/%Y", errors="coerce")

# Calculate today's date
today = pd.to_datetime("today")

# Calculate the aging days relative to today's date
df_status["Aging Days"] = df_status.apply(calculate_aging_days, axis=1)

# Apply the custom function to create a new "Status" column
df_status["Status"] = df_status["Aging Days"].apply(categorize_status)

# Merging the 2 files info together
df_merged = pd.merge(df_status,df_contact, on="Vendor", how="left")

# Sort columns 
column_order = ["S/N", "Vendor", "Item Bought", "Cost per unit", "Qty", "Total Cost", "Due date", "Payment Made", "Aging Days","Status", "Name", "Email address"]
df_merged = df_merged[column_order]

# Save the modified file
output_folder = r"C:\Users\Auklet\Dropbox\Cloud Projects\python email"
output_directory = os.path.join(output_folder, f"Update_Status_modified_{status_latest_date}.csv")
df_merged.to_csv(output_directory, index=False)
print("Modified file generated successfully!")

# Gather the vendors information
unique_vendors = df_merged["Vendor"].unique()

# Generate emails and send out
for vendor in unique_vendors:
    # Filter for the current vendor
    vendor_data = df_merged[(df_merged["Vendor"] == vendor)]

    # Filter data based on the "Status" column:Overdue and due in 7 days
    required_status = ["Overdue", "Due in < 7 days"]
    vendor_data = vendor_data[vendor_data["Status"].isin(required_status)]

    # Filter for the vendor email address
    recipient_email = vendor_data["Email address"].iloc[0]

    cc_list = "myboss@example.com; colleagues@example.com"
    item_table = generate_html_table(vendor_data)

    subject = f"Action Required: Outstanding Payment Notice ({vendor})"

    body = f"Dear {vendor},"
    body += f'''
    <p>We have noticed that the payment for the below-mentioned item(s) is/are still pending. Kindly make the payment for the below items. 
    For those items that are overdue, please make the payment by the stated date</p>
    <p>Your prompt attention to this matter would be greatly appreciated.</p>
    <p><strong>Payment by: <span style="color:red;">{due_date}</span></strong></p>
    '''
    body += item_table
    body += "<br>Thank you.<p>Best Regards,<br>XXX</p>"
    
    # for this example i have not added any attachment, but you could include the data in a csv and attach to the email.
    send_email(recipient_email,cc_list,subject, body)

print("Emails generated successfully!")
