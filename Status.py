from twilio.twiml.messaging_response import MessagingResponse
from google.oauth2.service_account import Credentials
from flask import Flask, request
from dotenv import load_dotenv
import gspread
import pandas as pd
import base64
import tempfile
import re
import os

# Load environment variables from .env
load_dotenv()

# Accessing the variables
TWILIO_ACCOUNT_SID = os.getenv("Twilio_Account_SID")
TWILIO_AUTH_TOKEN = os.getenv("Twilio_Auth_Token")
TWILIO_PHONE_NUMBER = os.getenv("Twilio_Phone_Number")

# Print to check if values are loaded (remove in production)
# print("Twilio SID:", TWILIO_ACCOUNT_SID)
# print("Twilio Auth Token:", TWILIO_AUTH_TOKEN)
# print("Twilio Phone Number:", TWILIO_PHONE_NUMBER)

# Step 1: Decode the base64 credentials
print(f"Env Variable Found: {os.getenv('Google_Sheets_Credentials') is not None}")
credentials_data = os.getenv('Google_Sheets_Credentials')
if credentials_data:
    credentials_json = base64.b64decode(credentials_data)

    # Create a temporary file to store the credentials
    with tempfile.NamedTemporaryFile(delete=False, mode='wb') as temp_file:
        temp_file.write(credentials_json)
        temp_file_path = temp_file.name

    # Step 2: Authenticate & Connect to Google Sheets
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    credentials = Credentials.from_service_account_file(temp_file_path, scopes=scope)
    client = gspread.authorize(credentials)
    
    # Clean up temporary file (optional)
    os.remove(temp_file_path)
else:
    print("❌ Error: Google Sheets credentials not found in environment variables.")

# Step 3: Open Google Sheet
google_sheets_url = os.getenv("google_sheets_url")
sheet = client.open_by_url(google_sheets_url)
print("✅ Successfully connected to Google Sheets!")

# Step 4: User Inputs
# print("Enter the status message (Press Enter twice to finish):")
# message = ""
# while True:
#     line = input()
#     if line == "":
#         break
#     message += line + "\n"
# 
# print("Received Message:\n", message)
app = Flask(__name__)
@app.route("/webhook", methods=["POST"])

def receive_whatsapp_message():
    # Webhook to receive WhatsApp messages from Twilio.
    message = request.values.get("Body", "").strip()  # Get message text from Twilio
    sender = request.values.get("From", "")  # Get sender's number
    
    print(f"📩 New message from {sender}: {message}")  # Debugging

    # Process message (extract info & update Google Sheets)
    response_text = process_message(message)

    # Send a reply
    resp = MessagingResponse()
    resp.message(response_text)
    return str(resp)

# Step 5: Define Official Status Mapping
status_mapping = {
    "PRESENT": "PRESENT",
    "ATTACH IN": "ATTACH IN",
    "DUTY": "DUTY",
    "UDO": "DUTY",
    "CDOS": "DUTY",
    "Guard Duty": "DUTY",
    "Guard Rest": "DUTY",
    "WFH": "WFH",
    "OUTSTATION": "OUTSTATION",
    "OS": "OUTSTATION",
    "CSE": "CSE",
    "AO": "AO",
    "LEAVE": "LEAVE",
    "OFF": "OFF",
    "RSI/RSO": "RSI/RSO",
    "RSI": "RSI/RSO",
    "RSO": "RSI/RSO",
    "MC": "MC",
    "MA": "MA"
}

# Step 6: Extract Fields
def extract_message(message):
    # Extract Status and Location (if in "Status:")
    status_match = re.search(r"Status:\s*([A-Z]+(?:\s+[A-Z]+)?)\s*(?:\s*@\s*(.+))?$", message, re.IGNORECASE | re.MULTILINE)
    raw_status = status_match.group(1).strip() if status_match else "Unknown"
    location = status_match.group(2).strip() if status_match and status_match.group(2) else ""

    # Convert status to official version
    status = status_mapping.get(raw_status.upper(), "Invalid")

    if status == "Invalid":
        print(f"❌ Error: '{raw_status}' is not a valid status.")
        exit()

    # Extract Names (Handles "R/Name" with or without ":" and same-line names)
    name_lines = []
    name_section = False

    for line in message.split("\n"):
        line = line.strip()
        
        # If "R/Name" is found, start capturing names
        match = re.match(r"R/Name:?\s*(.*)", line, re.IGNORECASE)
        if match:
            name_section = True
            names_in_line = match.group(1).strip()
            if names_in_line:  # If names are on the same line as "R/Name"
                name_lines.extend(names_in_line.split("\n"))  
            continue

        # Stop capturing names if "Dates:" is found
        elif re.match(r"Dates:?", line, re.IGNORECASE):
            name_section = False
            break  

        # If inside the "R/Name" section, collect names
        elif name_section:
            name_lines.append(line)

    # Remove the first word (rank) from each name
    names = [" ".join(name.split()[1:]) for name in name_lines if len(name.split()) > 1]

    # Extract Date and determine which sheets to update
    date_match = re.search(r"Dates?:\s*(\d{2}/\d{2}/\d{2,4})(?:\s*\(?(AM|PM)?\)?)?", message, re.IGNORECASE)
    date_text = date_match.group(1) if date_match else "Unknown"
    period = date_match.group(2) if date_match and date_match.group(2) else ""

    # Determine sheets to update
    sheets_to_update = []
    if period == "AM":
        sheets_to_update.append("AM")
    elif period == "PM":
        sheets_to_update.append("PM")
    else:
        sheets_to_update.extend(["AM", "PM"])  # If no period specified, update both

    # Night sheet updates only for specific statuses
    if status in ["DUTY", "CSE", "AO", "LEAVE", "OFF", "MC"]:
        sheets_to_update.append("NIGHT")

    # Extract Reason (if exists)
    reason_match = re.search(r"Reason:\s*(.+)", message, re.IGNORECASE)
    reason = reason_match.group(1).strip() if reason_match else ""

    # Extract Location (if provided separately)
    location_match = re.search(r"Location:\s*(.+)", message, re.IGNORECASE)
    if location_match:
        location = location_match.group(1).strip()

    # Output extracted values
    print("Extracted Status:", status)
    print("Extracted Location:", location)
    print("Extracted Names:", names)
    print("Extracted Date:", date_text)
    print("Extracted Reason:", reason)
    print("Sheets to update:", sheets_to_update)
    return status, location, names, date_text, reason, sheets_to_update

# Step 7: Update Google Sheets for each sheet
def update_sheet(status, location, names, date_text, reason, sheets_to_update):
    for sheet_name in sheets_to_update:
        worksheet = sheet.worksheet(sheet_name)
        data = worksheet.get_all_values()
        headers = data[1]  # Use second row as headers
        df = pd.DataFrame(data[2:], columns=headers)  # Data starts from the third row

        # Column indices
        try:
            status_col = headers.index("Status")
            date_col = headers.index("Date")
            remarks_col = headers.index("Remarks")
            location_col = headers.index("Location")
        except ValueError:
            print(f"❌ Error: Required columns missing in {sheet_name} sheet.")
            continue

        # Update each person's record
        for name in names:
            matching_rows = df[df["Name"].str.contains(name, case=False, na=False)].index.tolist()

            if not matching_rows:
                print(f"⚠️ No matching name found in {sheet_name} sheet for '{name}'")
                continue

            row_index = matching_rows[0] + 3  # Adjusting for header rows

            # Update the Google Sheet
            updates = [
                (f"{chr(65 + status_col)}{row_index}", [[status]]),
                (f"{chr(65 + date_col)}{row_index}", [[date_text]]),
                (f"{chr(65 + remarks_col)}{row_index}", [[reason]]),
                (f"{chr(65 + location_col)}{row_index}", [[location]])
            ]

            for cell, value in updates:
                worksheet.update(range_name=cell, values=value)

            print(f"✅ Successfully updated {name}'s record in {sheet_name} sheet (Row {row_index})")

    print("✅ All updates completed!")

def process_message(message):
    status, location, names, date_text, reason, sheets_to_update = extract_message(message)
    update_sheet(status, location, names, date_text, reason, sheets_to_update)

if __name__ == "__main__":
    from waitress import serve  # More efficient than Flask's built-in server
    port = int(os.getenv("PORT", 8080))
    serve(app, host="0.0.0.0", port=port)