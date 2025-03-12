# from twilio.twiml.messaging_response import MessagingResponse
from google.oauth2.service_account import Credentials
from flask import Flask, request
from dotenv import load_dotenv
import gspread
import pandas as pd
import base64
import tempfile
import re
import os

# from telegram import ParseMode
from telegram.ext import Application, CommandHandler, CallbackContext


# Load environment variables from .env
load_dotenv()

# Accessing the variables
# TWILIO_ACCOUNT_SID = os.getenv("Twilio_Account_SID")
# TWILIO_AUTH_TOKEN = os.getenv("Twilio_Auth_Token")
# TWILIO_PHONE_NUMBER = os.getenv("Twilio_Phone_Number")
# TWILIO_SERVICE_SID = os.getenv("Twilio_Service_SID")

# Telegram Bot Token
TELEGRAM_TOKEN = os.getenv('Telegram_Token')
print("Telegram Token: ", os.getenv('Telegram_Token'))

if not TELEGRAM_TOKEN:
    raise ValueError("Telegram Token is missing from the environment variables!")
# bot = Bot(token=TELEGRAM_TOKEN)

application = Application.builder().token(TELEGRAM_TOKEN).build()

# Example command handler, like /start
async def start(update, context):
    await update.message.reply_text('Hello! I\'m your bot!')

# Adding the handler to the application
application.add_handler(CommandHandler("start", start))

# Run the bot
application.run_polling()



# # Step 1: Decode the base64 credentials
# print(f"Env Variable Found: {os.getenv('Google_Sheets_Credentials') is not None}")
# credentials_data = os.getenv('Google_Sheets_Credentials')
# if credentials_data:
#     credentials_json = base64.b64decode(credentials_data)

#     # Create a temporary file to store the credentials
#     with tempfile.NamedTemporaryFile(delete=False, mode='wb') as temp_file:
#         temp_file.write(credentials_json)
#         temp_file_path = temp_file.name

#     # Step 2: Authenticate & Connect to Google Sheets
#     scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
#     credentials = Credentials.from_service_account_file(temp_file_path, scopes=scope)
#     client = gspread.authorize(credentials)
    
#     # Clean up temporary file (optional)
#     os.remove(temp_file_path)
# else:
#     print("❌ Error: Google Sheets credentials not found in environment variables.")

# # Step 3: Open Google Sheet
# google_sheets_url = os.getenv("google_sheets_url")
# sheet = client.open_by_url(google_sheets_url)
# print("✅ Successfully connected to Google Sheets!")

# # Step 4: User Inputs
# app = Flask(__name__)

# # Store pending updates (key: sender number, value: extracted info)
# updates = {}

# # Route to check if the server is running
# @app.route("/", methods=["GET"])
# def home():
#     return "Flask server is running!", 200

# # Route to handle incoming webhook messages from Telegram
# @app.route("/webhook", methods=["POST"])
# def webhook():
#     # Webhook to receive messages from Telegram.
#     # message = request.values.get("Body", "").strip()
#     # sender = request.values.get("From", "")
#     message = request.json
#     sender = message['message']['from']['id']
#     text = message['message']['text'].strip()

#     print(f"📩 New message from {sender}: \n{text}")  # Debugging


#     # # Check if user is replying with a button
#     # interactive_response = request.values.get("InteractiveResponseId")

#     # if interactive_response:
#     #     return handle_interactive_response(sender, interactive_response)

#     # Process the message and extract details
#     status, location, names, date_text, reason, sheets_to_update = extract_message(text)


#     # Store the extracted info in dictionary updates
#     updates[sender] = {
#         "status": status, "location": location, "names": names,
#         "date_text": date_text, "reason": reason, "sheets_to_update": sheets_to_update
#     }

#     # Send quick reply buttons to confirm update
#     send_confirmation_message(sender, updates[sender])

#     # # Send multiple messages
#     # response = MessagingResponse()
#     # response.message(f"✅ Received your message: {message}\n"
#     #              f"📌 Status: {status}\n"
#     #              f"📍 Location: {location}\n"
#     #              f"👥 Names: {', '.join(names) if names else 'None'}\n"
#     #              f"📅 Dates: {date_text}\n"
#     #              f"📄 Reason: {reason}")

#     # # Update Google Sheets
#     # complete = update_sheet(status, location, names, date_text, reason, sheets_to_update)

#     # if complete:
#     #     response.message("✅ All updates completed!")
#     # else:
#     #     response.message("❌ Error: Check logs for issue...")

#     # return str(response)
#     return "OK", 200

# # Step 5: Define Official Status Mapping
# status_mapping = {
#     "PRESENT": "PRESENT",
#     "ATTACH IN": "ATTACH IN",
#     "DUTY": "DUTY",
#     "UDO": "DUTY",
#     "CDOS": "DUTY",
#     "GUARD": "DUTY",
#     "WFH": "WFH",
#     "OUTSTATION": "OUTSTATION",
#     "BLOOD DONATION": "OUTSTATION",
#     "OS": "OUTSTATION",
#     "CSE": "CSE",
#     "AO": "AO",
#     "LEAVE": "LEAVE",
#     "OFF": "OFF",
#     "RSI/RSO": "RSI/RSO",
#     "RSI": "RSI/RSO",
#     "RSO": "RSI/RSO",
#     "MC": "MC",
#     "MA": "MA"
# }

# # Step 6: Extract Fields
# def extract_message(message):
#     # Extract Status and Location (if in "Status:")
#     status_match = re.search(r"Status:\s*([A-Z]+(?:\s+[A-Z]+)?)\s*(?:\s*@\s*(.+))?$", message, re.IGNORECASE | re.MULTILINE)
#     raw_status = status_match.group(1).strip() if status_match else "Unknown"
#     location = status_match.group(2).strip() if status_match and status_match.group(2) else ""

#     # Convert status to official version
#     # status = status_mapping.get(raw_status.upper(), "Invalid")

#     for keyword in status_mapping.keys():
#         if keyword.upper() in raw_status.upper():  # Case-insensitive matching
#             status = status_mapping[keyword]  # Return mapped status
#             break
#         else:
#             status = "Invalid"

#     if status == "Invalid":
#         print(f"❌ Error: '{raw_status}' is not a valid status.")
#         return "❌ Invalid status detected.", 400

#     # Extract Names (Handles "R/Name" with or without ":" and same-line names)
#     name_lines = []
#     name_section = False

#     for line in message.split("\n"):
#         line = line.strip()
        
#         # If "R/Name" is found, start capturing names
#         match = re.match(r"R/Name:?\s*(.*)", line, re.IGNORECASE)
#         if match:
#             name_section = True
#             names_in_line = match.group(1).strip()
#             if names_in_line:  # If names are on the same line as "R/Name"
#                 name_lines.extend(names_in_line.split("\n"))  
#             continue

#         # Stop capturing names if "Dates:" is found
#         elif re.match(r"Dates?\s*:?", line, re.IGNORECASE):
#             name_section = False
#             break  

#         # If inside the "R/Name" section, collect names
#         elif name_section:
#             name_lines.append(line)

#     # Remove the first word (rank) from each name
#     names = [" ".join(name.split()[1:]) for name in name_lines if len(name.split()) > 1]

#     # Extract Date and determine which sheets to update
#     # date_match = re.search(r"Dates?\s*:?\s*(\d{2}/\d{2}/\d{2,4})(?:\s*\(?(AM|PM)?\)?)?", message, re.IGNORECASE)
#     # date_text = date_match.group(1) if date_match else ""
#     # period = date_match.group(2) if date_match and date_match.group(2) else ""

#     date_match = re.search(r"Dates?\s*:?\s*(.+)", message, re.IGNORECASE)
#     date_text = date_match.group(1).strip() if date_match else ""
#     period = ""
#     if "(AM)" in date_text.upper():
#         period = "AM"
#     elif "(PM)" in date_text.upper():
#         period = "PM"


#     # Determine sheets to update
#     sheets_to_update = []
#     if period == "AM":
#         sheets_to_update.append("AM")
#     elif period == "PM":
#         sheets_to_update.append("PM")
#     else:
#         sheets_to_update.extend(["AM", "PM"])  # If no period specified, update both

#     # Night sheet updates only for specific statuses
#     if status in ["DUTY", "CSE", "AO", "LEAVE", "OFF", "MC"]:
#         sheets_to_update.append("NIGHT")

#     # Extract Reason (if exists)
#     reason_match = re.search(r"Reason:\s*(.+)", message, re.IGNORECASE)
#     reason = reason_match.group(1).strip() if reason_match else ""

#     # Extract Location (if provided separately)
#     location_match = re.search(r"Location:\s*(.+)", message, re.IGNORECASE)
#     if location_match:
#         location = location_match.group(1).strip()

#     # Output extracted values
#     print("Extracted Status:", status)
#     print("Extracted Location:", location)
#     print("Extracted Names:", names)
#     print("Extracted Date:", date_text)
#     print("Extracted Reason:", reason)
#     print("Sheets to update:", sheets_to_update)
#     return status, location, names, date_text, reason, sheets_to_update

# # Step 7: Update Google Sheets for each sheet
# def update_sheet(status, location, names, date_text, reason, sheets_to_update):
#     for sheet_name in sheets_to_update:
#         worksheet = sheet.worksheet(sheet_name)
#         data = worksheet.get_all_values()
#         headers = data[1]  # Use second row as headers
#         df = pd.DataFrame(data[2:], columns=headers)  # Data starts from the third row

#         # Column indices
#         try:
#             status_col = headers.index("Status")
#             date_col = headers.index("Date")
#             remarks_col = headers.index("Remarks")
#             location_col = headers.index("Location")
#         except ValueError:
#             print(f"❌ Error: Required columns missing in {sheet_name} sheet.")
#             continue

#         # Update each person's record
#         for name in names:
#             matching_rows = df[df["Name"].str.contains(name, case=False, na=False)].index.tolist()

#             if not matching_rows:
#                 print(f"⚠️ No matching name found in {sheet_name} sheet for '{name}'")
#                 continue

#             row_index = matching_rows[0] + 3  # Adjusting for header rows

#             # Update the Google Sheet
#             updates = [
#                 (f"{chr(65 + status_col)}{row_index}", [[status]]),
#                 (f"{chr(65 + date_col)}{row_index}", [[date_text]]),
#                 (f"{chr(65 + remarks_col)}{row_index}", [[reason]]),
#                 (f"{chr(65 + location_col)}{row_index}", [[location]])
#             ]

#             for cell, value in updates:
#                 worksheet.update(range_name=cell, values=value)

#             print(f"✅ Successfully updated {name}'s record in {sheet_name} sheet (Row {row_index})")

#     print("✅ All updates completed!")
#     return True

# # Step 8: Confirm status update
# def send_confirmation_message(chat_id, extracted_info):
#     # Sends a message with confirmation buttons to confirm or cancel update.
#     # url = "https://api.twilio.com/2010-04-01/Accounts/YOUR_ACCOUNT_SID/Messages.json"

#     # headers = {"Content-Type": "application/x-www-form-urlencoded"}

#     # data = {
#     #     "From": f"whatsapp:+{TWILIO_PHONE_NUMBER}",
#     #     "To": to,
#     #     "MessagingServiceSid": f"{TWILIO_SERVICE_SID}",
#     #     "Content-Type": "application/json",
#     #     "InteractiveMessage": json.dumps({
#     #         "type": "button",
#     #         "body": {
#     #             "text": f"✅ Recieved message\n"
#     #                     f"📌 Status: {extracted_info['status']}\n"
#     #                     f"📍 Location: {extracted_info['location']}\n"
#     #                     f"👥 Names: {', '.join(extracted_info['names'])}\n"
#     #                     f"📅 Dates: {extracted_info['date_text']}\n"
#     #                     f"📄 Reason: {extracted_info['reason']}\n\n"
#     #                     "Is this status correct?"
#     #         },
#     #         "action": {
#     #             "buttons": [
#     #                 {
#     #                     "type": "reply",
#     #                     "reply": {
#     #                         "id": "confirm_update",
#     #                         "title": "✅ Confirm"
#     #                     }
#     #                 },
#     #                 {
#     #                     "type": "reply",
#     #                     "reply": {
#     #                         "id": "cancel_update",
#     #                         "title": "❌ Cancel"
#     #                     }
#     #                 }
#     #             ]
#     #         }
#     #     })
#     # }

#     # auth = (TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)

#     # response = requests.post(url, data=data, headers=headers, auth=auth)
#     # print(response.json())  # Debugging

#     """ Sends a message with confirmation buttons to Telegram. """
#     text = (f"📌 Status: {extracted_info['status']}\n"
#             f"📍 Location: {extracted_info['location']}\n"
#             f"👥 Names: {', '.join(extracted_info['names'])}\n"
#             f"📅 Dates: {extracted_info['date_text']}\n"
#             f"📄 Reason: {extracted_info['reason']}\n\n"
#             "Do you want to update this status?")

#     keyboard = [[
#         {'text': '✅ Confirm', 'callback_data': 'confirm_update'},
#         {'text': '❌ Cancel', 'callback_data': 'cancel_update'}
#     ]]

#     bot.send_message(chat_id, text, reply_markup={'inline_keyboard': keyboard}, parse_mode=ParseMode.MARKDOWN)

# def handle_interactive_response(chat_id, callback_data):
#     # Handles user responses to interactive buttons.
#     # response = MessagingResponse()

#     # if sender not in updates:
#     #     response.message("⚠ No pending update found.")
#     #     return str(response)

#     # if button_id == "confirm_update":
#     #     # Retrieve pending update and update Google Sheets
#     #     extracted_info = updates.pop(sender)
#     #     complete = update_sheet(
#     #         extracted_info["status"], extracted_info["location"], extracted_info["names"],
#     #         extracted_info["date_text"], extracted_info["reason"], extracted_info["sheets_to_update"]
#     #     )

#     #     if complete:
#     #         response.message("✅ Status update successful!")
#     #     else:
#     #         response.message("❌ Error updating the sheet. Please try again.")

#     # elif button_id == "cancel_update":
#     #     updates.pop(sender, None)
#     #     response.message("❌ Update cancelled.")

#     # return str(response)

#     if chat_id not in updates:
#         bot.send_message(chat_id, "⚠ No pending update found.")
#         return

#     if callback_data == "confirm_update":
#         extracted_info = updates.pop(chat_id)
#         complete = update_sheet(
#             extracted_info["status"], extracted_info["location"], extracted_info["names"],
#             extracted_info["date_text"], extracted_info["reason"], extracted_info["sheets_to_update"]
#         )

#         if complete:
#             bot.send_message(chat_id, "✅ Status update successful!")
#         else:
#             bot.send_message(chat_id, "❌ Error updating the sheet. Please try again.")
#     elif callback_data == "cancel_update":
#         updates.pop(chat_id, None)
#         bot.send_message(chat_id, "❌ Update cancelled.")

# if __name__ == "__main__":
#     from waitress import serve  # More efficient than Flask's built-in server
#     port = int(os.getenv("PORT", 8080))
#     serve(app, host="0.0.0.0", port=port)'