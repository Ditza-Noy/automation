import pandas as pd
import pywhatkit
import time
import os

# --- Configuration ---
# The exact name of your original Excel file
contacts_filename = 'רשימת דיירים .xlsx'

# The specific sheet name to read from
sheet_name = 'רשימת דיירים ובעלי דירות'

# The messages file created by you (Columns: Apartment, Message, Target)
messages_filename = 'messages.xlsx'

print("Starting the automation process...")

try:
    # 1. Check if files exist
    if not os.path.exists(contacts_filename):
        print(f"Error: The file '{contacts_filename}' was not found.")
        print("Please ensure the file is in the same folder as this script.")
        exit()
        
    if not os.path.exists(messages_filename):
        print(f"Error: The file '{messages_filename}' was not found.")
        exit()

    # 2. Load the data
    print(f"Reading contacts from '{contacts_filename}' (Sheet: {sheet_name})...")
    
    # Reading the specific sheet from the Excel file
    # engine='openpyxl' is required for reading .xlsx files
    df_contacts = pd.read_excel(contacts_filename, sheet_name=sheet_name, engine='openpyxl')
    
    # Reading the messages file
    df_messages = pd.read_excel(messages_filename)

    # 3. Merge data based on 'Apartment' column
    # This combines the message with the phone numbers
    df_final = pd.merge(df_messages, df_contacts, on='Apartment', how='left')
    
    total_messages = len(df_final)
    print(f"Found {total_messages} messages to process.")
    print("Please do not touch the mouse or keyboard during execution.")
    
    # 4. Processing Loop
    for index, row in df_final.iterrows():
        apartment_num = row['Apartment']
        
        # --- Logic: Choose Target (Tenant or Owner) ---
        target_val = str(row.get('Target', '')).strip()
        
        # Check if the target is the Owner (supports Hebrew 'בעל' or English 'Owner')
        if 'בעל' in target_val or 'Owner' in target_val:
            raw_phone = row['Owner_Phone']
            target_type = "Owner"
        else:
            # Default is Tenant
            raw_phone = row['Tenant_Phone']
            target_type = "Tenant"

        # Check if phone number exists
        if pd.isna(raw_phone):
            print(f"Skipping Apartment {apartment_num} ({target_type}) - No phone number found.")
            continue

        # Format the phone number (Ensure it starts with +)
        phone = str(raw_phone)
        if not phone.startswith('+'):
            phone = '+' + phone
            
        message = row['Message']
        
        print(f"Sending to Apartment {apartment_num} -> Target: {target_type} (Phone: {phone})...")
        
        try:
            # Send the message
            pywhatkit.sendwhatmsg_instantly(
                phone, 
                message, 
                wait_time=12, 
                tab_close=True, 
                close_time=3
            )
            
            # Safety delay between messages
            time.sleep(8) 
            
        except Exception as send_error:
            print(f"Failed to send to Apartment {apartment_num}: {send_error}")

    print("Process completed successfully! All messages sent.")

except Exception as e:
    print(f"An unexpected error occurred: {e}")