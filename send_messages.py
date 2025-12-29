import pandas as pd
import pywhatkit
import time
import os
import pyautogui # Added this library to control the keyboard

# --- Configuration ---
contacts_filename = 'רשימת דיירים .xlsx'
sheet_name = 'רשימת דיירים ובעלי דירות'
messages_filename = 'messages.xlsx'

print("Starting the automation process...")

try:
    # 1. Check if files exist
    if not os.path.exists(contacts_filename):
        print(f"Error: The file '{contacts_filename}' was not found.")
        exit()
        
    if not os.path.exists(messages_filename):
        print(f"Error: The file '{messages_filename}' was not found.")
        exit()

    # 2. Load the data
    print(f"Reading contacts from '{contacts_filename}'...")
    df_contacts = pd.read_excel(contacts_filename, sheet_name=sheet_name, engine='openpyxl')
    df_messages = pd.read_excel(messages_filename)

    # 3. Merge data
    df_final = pd.merge(df_messages, df_contacts, on='Apartment', how='left')
    
    print(f"Found {len(df_final)} messages to process.")
    print("IMPORTANT: Do not touch the mouse or keyboard!")
    
    # 4. Processing Loop
    for index, row in df_final.iterrows():
        apartment_num = row['Apartment']
        
        # Target Logic (Tenant vs Owner)
        target_val = str(row.get('Target', '')).strip()
        if 'בעל' in target_val or 'Owner' in target_val:
            raw_phone = row['Owner_Phone']
            target_type = "Owner"
        else:
            raw_phone = row['Tenant_Phone']
            target_type = "Tenant"

        if pd.isna(raw_phone):
            print(f"Skipping Apt {apartment_num}: No phone number.")
            continue

        phone = str(raw_phone)
        if not phone.startswith('+'):
            phone = '+' + phone
            
        message = row['Message']
        
        print(f"Sending to Apt {apartment_num} ({target_type})...")
        
        try:
            # Open WhatsApp and type the message
            # Increased wait_time to 15 to ensure page loads fully
            pywhatkit.sendwhatmsg_instantly(
                phone, 
                message, 
                wait_time=15, 
                tab_close=False # We will close it manually later if needed
            )
            
            # FORCE SEND FIX:
            time.sleep(2) # Wait for text to appear
            pyautogui.press('enter') # Press Enter to send
            time.sleep(1)
            pyautogui.press('enter') # Press again just in case (sometimes needed)
            
            # Close the tab (Optional - currently disabled to be safe)
            # pyautogui.hotkey('ctrl', 'w') 
            
            print("Message sent.")
            
            # Wait before next message
            time.sleep(6) 
            
        except Exception as send_error:
            print(f"Failed to send: {send_error}")

    print("Process completed.")

except Exception as e:
    print(f"An unexpected error occurred: {e}")