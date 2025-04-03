from traceback import print_tb
from telegram.ext import Application, CommandHandler, MessageHandler, filters, CallbackQueryHandler, ContextTypes
from telegram.constants import ParseMode
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext._contexttypes import ContextTypes
from contextlib import asynccontextmanager
from google.oauth2.service_account import Credentials
from fastapi import FastAPI, Request, Response
from dotenv import load_dotenv
from http import HTTPStatus
import pandas as pd
import tempfile
import gspread
import base64
import re
import os
from apscheduler.schedulers.background import BackgroundScheduler
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import asyncio
# Load environment variables from .env
load_dotenv()

# Telegram Bot Token
TELEGRAM_TOKEN = os.getenv('Telegram_Token')
CHAT_ID = os.getenv('Chat_ID')
# print("Telegram Token: ", os.getenv('Telegram_Token'))

if not TELEGRAM_TOKEN:
    raise ValueError("Telegram Token is missing from the environment variables!")

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
google_sheets_url = os.getenv("google_sheets_url") # AI Sheet
real_google_sheets_url = os.getenv("real_google_sheets_url")
sheet = client.open_by_url(real_google_sheets_url) # Change to toggle
informal_google_sheets_url = os.getenv("informal_google_sheet_url") # AI Sheet
real_informal_google_sheets_url = os.getenv("real_informal_google_sheets_url")
informal_sheet = client.open_by_url(real_informal_google_sheets_url) # Change to toggle
print("✅ Successfully connected to Google Sheets!")

# Step 4: Building the bot
ptb = (
    Application.builder()
    .updater(None)
    .token(TELEGRAM_TOKEN)  # replace <your-bot-token>
    .read_timeout(7)
    .get_updates_read_timeout(42)
    .build()
)

# Send message when bot starts
async def send_startup_message():
    # Replace with the chat ID where you want to send the message
    chat_id = CHAT_ID  # Can be your own chat ID or a group chat ID
    await ptb.bot.send_message(chat_id, "Startup complete!")
    asyncio.create_task(start_scheduler())
    print("✅ Scheduler started")

# /Start command handler
async def start(update: Update, _: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Starting...")
ptb.add_handler(CommandHandler("start", start))

# /id Get chat id
async def get_chat_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Reply with the user's chat ID
    chat_id = update.message.chat_id
    await update.message.reply_text(f"Your chat ID is: {chat_id}")
ptb.add_handler(CommandHandler("id", get_chat_id))

# /check Manually run status check
async def check_status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    telegram_message = await ptb.bot.send_message(chat_id=CHAT_ID, text="🔄 Checking status...")
    message = await check_and_update_status()
    await telegram_message.edit_text(message)
ptb.add_handler(CommandHandler("check", check_status))

# Function for other functions to send Telegram message
async def send_telegram_message(message: str):
    await ptb.bot.send_message(chat_id=CHAT_ID, text=message)

@asynccontextmanager
async def lifespan(_: FastAPI):
    await ptb.bot.deleteWebhook()  # Ensure webhook is reset
    await ptb.bot.setWebhook("https://updatestatus-production.up.railway.app/webhook") # replace <your-webhook-url>
    async with ptb:
        await send_startup_message()
        await ptb.start()
        yield
        await ptb.stop()

app = FastAPI(lifespan=lifespan)

@app.get("/ping")
async def ping():
    return {"message": "pong"}

@app.post("/webhook")
async def process_update(request: Request):
    req = await request.json()
    update = Update.de_json(req, ptb.bot)
    await ptb.process_update(update)
    return Response(status_code=HTTPStatus.OK)

# Message handler for processing status updates
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message.text.strip()
    sender = update.message.from_user.id
 
    print(f"📩 New message from {sender}: \n{message}") # Debugging

    # Process the message and extract details
    data = extract_message(message)
    status, informal_status, location, names, date_text, reason, sheets_to_update, informal_sheets_to_update = data["status"], data["informal_status"], data["location"], data["names"], data["date_text"], data["reason"], data["sheets_to_update"], data["informal_sheets_to_update"]

    # Confirmation button
    keyboard = [
        [
            InlineKeyboardButton("✅ Confirm", callback_data="confirm"),
            InlineKeyboardButton("❌ Cancel", callback_data="cancel")
        ]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)

    # Send multiple messages
    response = (f"✅ Status Update Recieved\n"
                    f"📌 Status: {status}\n"
                    f"📌 Informal Status: {informal_status}\n"
                    f"📍 Location: {location}\n"
                    f"👥 Names: {', '.join(names) if names else 'None'}\n"
                    f"📅 Dates: {date_text}\n"
                    f"📄 Reason: {reason}\n"
                    f"📄 Sheets: {sheets_to_update}\n"
                    f"📄 Informal Sheets: {informal_sheets_to_update}\n")
    await update.message.reply_text(response, reply_markup=reply_markup, parse_mode="Markdown"), 

    # Wait for user to confirm update
    context.user_data["status_data"] = (status, informal_status, location, names, date_text, reason, sheets_to_update, informal_sheets_to_update)
ptb.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))  # Handles all text messages

async def handle_confirmation(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    data =  query.data
    print("Waiting for response...")

    # Remove the buttons
    await query.edit_message_reply_markup(reply_markup=None)

    if data == "cancel":
        await query.message.reply_text("❌ Status update cancelled.")
        return
    loading = await query.message.reply_text("🔄 Updating status...")

    # Get data and update sheet
    status, informal_status, location, names, date_text, reason, sheets_to_update, informal_sheets_to_update = context.user_data.pop('status_data', (None,)*8)

    if not status:
        print("Error: No data found in context.")
        return
    
    # Update excel sheet
    complete = await update_sheet(status, location, names, date_text, reason, sheets_to_update)
    success = await update_informal_sheet(informal_status, names, date_text, informal_sheets_to_update)
    if complete and success:
        await loading.edit_text("✅ All updates completed!")
    else:
        await loading.edit_text("⚠️ Error: Check logs for issue...")
ptb.add_handler(CallbackQueryHandler(handle_confirmation))

# Step 5: Define Official and Informal Status Mapping
official_status_mapping = {
    "ATTACH IN": "ATTACH IN",
    "DUTY": "DUTY",
    "UDO": "DUTY",
    "CDOS": "DUTY",
    "GUARD": "DUTY",
    "REST": "DUTY",
    "WFH": "WFH",
    "OUTSTATION": "OUTSTATION",
    "BLOOD DONATION": "OUTSTATION",
    "OS": "OUTSTATION",
    "CSE": "CSE",
    "COURSE": "CSE",
    "AO": "AO",
    "OL": "LEAVE",
    "OVERSEAS": "LEAVE",
    "LEAVE": "LEAVE",
    "OFF": "OFF",
    "RSI": "RSI/RSO",
    "RSO": "RSI/RSO",
    "MC": "MC",
    "MA": "MA"
}

informal_status_mapping = {
    "GD": "GD",
    "DUTY": "GD",
    "UDO": "UDO",
    "CDOS": "CDOS",
    "GR": "GR",
    "REST": "GR",
    "OUTSTATION": "OS",
    "BLOOD": "OS",
    "DONATION": "OS",
    "OS": "OS",
    "CSE": "C",
    "COURSE": "C",
    "AO": "AO",
    "ATTACH": "AO",
    "OL": "OL",
    "OVERSEAS": "OL",
    "LEAVE": "L",
    "OFF": "O",
    "RSI": "RSI",
    "RSO": "RSO",
    "MC": "MC",
    "MA": "MA",
    "XWB": "XWB",
    "XTW": "XTW",
    "PRESENT": "1"
}

# Step 6: Extract information
def extract_message(message):
    lines = message.split("\n")

    # Extract Status and Location (if in "Status:")
    status_match = re.search(r"Status\s*:?\s*(.+?)\s*(?:@\s*(.+))?$", message, re.IGNORECASE | re.MULTILINE)
    raw_status = status_match.group(1).strip() if status_match else "Unknown"
    location = status_match.group(2).strip() if status_match and status_match.group(2) else ""

    # Convert status to official version
    status = "Invalid"
    for keyword in official_status_mapping.keys():
        if keyword.upper() in raw_status.upper():  # Case-insensitive matching
            status = official_status_mapping[keyword]  # Return mapped status
            break
    # Convert status to informal version
    for keyword in informal_status_mapping.keys():
        if keyword.upper() in raw_status.upper():  # Case-insensitive matching
            informal_status = informal_status_mapping[keyword]  # Return mapped status
            break

    if status == "Invalid":
        print(f"⚠️ Error: '{raw_status}' is not a valid status.")

    # Extract Names (Handles "R/Name" with or without ":" and same-line names)
    name_lines = []
    name_section = False

    for line in lines:
        line = line.strip()
        
        # If "R/Name" is found, start capturing names
        match = re.match(r"R/Names?\s*:?\s*(.+)", line, re.IGNORECASE)
        if match:
            name_section = True
            names_in_line = match.group(1).strip()
            if names_in_line:  # If names are on the same line as "R/Name"
                name_lines.extend(names_in_line.split("\n"))  
            continue

        # Stop capturing names if "Dates:" is found
        elif re.match(r"Dates?\s*:?", line, re.IGNORECASE):
            name_section = False
            break  

        # If inside the "R/Name" section, collect names
        elif name_section:
            name_lines.append(line)

    # Remove the first word (rank) from each name
    names = [" ".join(name.split()[1:]) for name in name_lines if len(name.split()) > 1]

    # Extract Dates
    date_match = re.search(r"Dates?\s*:?\s*(.+)", message, re.IGNORECASE)
    date_text = date_match.group(1).strip() if date_match else ""

    # Convert date to fixed format (DD/MM/YY)
    # Regular expression for 6-digit (DDMMYY) dates
    six_digit_pattern = r"\b(\d{2})(\d{2})(\d{2})\b"

    # Check if it's a range (date-date or date - date)
    sheets_to_update, informal_sheets_to_update = [], []
    informal_sheet_name = datetime.now().strftime("%b %y")    
    if "-" in date_text:
        # Normalize spaces around "-" and split the range
        start_date, end_date = [d.strip() for d in date_text.split("-")]

        # Convert individual dates
        start_date = re.sub(six_digit_pattern, r"\1/\2/\3", start_date)
        end_date = re.sub(six_digit_pattern, r"\1/\2/\3", end_date)
        date_text = f"{start_date} - {end_date}"

        # Determine AM or PM from both start and end dates
        start_am, start_pm = "(AM)" in start_date.upper(), "(PM)" in start_date.upper()
        end_am, end_pm = "(AM)" in end_date.upper(), "(PM)" in end_date.upper()

        # date_text is a range, all sheets need to be updated
        sheets_to_update.extend(["AM", "PM"])
        informal_sheets_to_update.extend([f"{informal_sheet_name} (AM)", f"{informal_sheet_name} (PM)"])        
    else:
        # Convert single date
        date_text = re.sub(six_digit_pattern, r"\1/\2/\3", date_text)
        if "(AM)" in raw_status.upper():
            date_text += " (AM)"
        if "(PM)" in raw_status.upper():
            date_text += " (PM)"

        # Determine AM or PM
        start_am, start_pm = "(AM)" in date_text.upper(), "(PM)" in date_text.upper()
        end_am, end_pm = False, False

        # Determine sheets to update
        if start_am or end_am:
            sheets_to_update.append("AM")
            informal_sheets_to_update.append(f"{informal_sheet_name} (AM)")
        if start_pm or end_pm:
            sheets_to_update.append("PM")
            informal_sheets_to_update.append(f"{informal_sheet_name} (PM)")
        if len(sheets_to_update) == 0:
            sheets_to_update.extend(["AM", "PM"])
            informal_sheets_to_update.extend([f"{informal_sheet_name} (AM)", f"{informal_sheet_name} (PM)"])

    # Night sheet updates only for specific statuses
    if status in ["DUTY", "CSE", "AO", "LEAVE", "OFF", "MC"] and len(sheets_to_update) == 2:
        sheets_to_update.append("NIGHT")

    # Extract Location and Reason (if provided separately)
    reason = "" # Location already has a check
    for line in lines:
        location_match = re.match(r"Locations?\s*:?\s*(.*)", line, re.IGNORECASE)
        if location_match:
            location = location_match.group(1).strip()

        reason_match = re.match(r"Reasons?\s*:?\s*(.*)", line, re.IGNORECASE)
        if reason_match:
            reason = reason_match.group(1).strip()

    # Output extracted values
    # print("Extracted Raw Status:", raw_status)
    print("Extracted Status:", status)
    print("Extracted Informal Status:", informal_status)
    print("Extracted Names:", names)
    print("Extracted Date:", date_text)
    print("Extracted Location:", location)
    print("Extracted Reason:", reason)
    print("Sheets to update:", sheets_to_update)
    print("Informal Sheets to update:", informal_sheets_to_update)
    return {
        "status": status,
        "informal_status": informal_status,
        "names": names,
        "date_text": date_text,
        "location": location,
        "reason": reason,
        "sheets_to_update": sheets_to_update,
        "informal_sheets_to_update": informal_sheets_to_update
    }

def extract_days(date_text):
    # Regex pattern to match date format and extract the day (the first two digits)
    day_pattern = r'(\d{2})/\d{2}/\d{2}'    # Only extracts the day (\d{2})

    # Search for all matches of the day pattern
    days = re.findall(day_pattern, date_text)

    if len(days) == 1:
        # If there is only one day, return that day
        return [str(int(days[0]))]
    elif len(days) == 2:
        # If there are two dates, generate the range of days
        start_day = int(days[0])
        end_day = int(days[1])
        
        # Generate all days within the range
        day_list = []
        current_day = start_day
        while current_day <= end_day:
            day_list.append(str(current_day))   # Append the day as a string
            current_day += 1                   # Increment by one day
        return day_list
    else:
        # In case of an unexpected format, return empty list
        return []

def get_column_letter(index):
    # Convert a 0-based column index to Excel-style column letters.
    letters = ""
    while index >= 0:
        letters = chr(index % 26 + 65) + letters
        index = index // 26 - 1
    return letters

# Step 7: Update Google Sheets for each sheet
async def update_sheet(status, location, names, date_text, reason, sheets_to_update):
    success, message = True, ""

    for sheet_name in sheets_to_update:
        worksheet = sheet.worksheet(sheet_name)
        data = worksheet.get_all_values()
        headers = data[1]  # Use second row as headers
        df = pd.DataFrame(data[2:], columns=headers)  # Data starts from the third row
        print(f"Accessing sheet: '{sheet_name}'")

        # Normalize headers by stripping leading/trailing whitespace
        formatted_headers = [header.strip() for header in headers]
        # Column indices
        try:
            status_col = formatted_headers.index("Status")
            date_col = formatted_headers.index("Date")
            remarks_col = formatted_headers.index("Remarks")
            location_col = formatted_headers.index("Location")
        except ValueError:
            success = False
            print(f"⚠️ Error: Required columns missing in {sheet_name} sheet.")
            continue

        # Collect all updates in a batch
        updates = []

        # Update each person's record
        for name in names:
            # Split the name into parts (tokens)
            name_parts = name.split()

            # Attempt to match each part of the name against the Excel sheet
            for part in name_parts:
                print(part)
                matching_rows = df[df["Name"].str.contains(part, case=False, na=False)].index.tolist()

                # Check if there's exactly one match (to confidently identify the person)
                if len(matching_rows) == 1:
                    print(f"✅ Match found: '{name}' matched with row index {matching_rows[0] + 3}")
                    break
                elif len(matching_rows) == 0:
                    print(f"⚠️ No matching name found in {sheet_name} sheet for '{name}'")
                else:
                    # Ambiguous matches (multiple rows match the name tokens)
                    print(f"⚠️ Multiple matches found for '{part}' in {sheet_name} sheet: {matching_rows}")

            if not matching_rows or status == "Invalid":
                success = False
                print(f"⚠️ No matching name found in {sheet_name} sheet for '{name}'" if not matching_rows else f"⚠️ Error: Status {status} for {name} is not valid.")
                continue

            row_index = matching_rows[0] + 3  # Adjusting for header rows

            # Update the Google Sheet
            updates.extend([
                {"range": f"{chr(65 + status_col)}{row_index}", "values": [[status]]},
                {"range": f"{chr(65 + date_col)}{row_index}", "values": [[date_text]]},
                {"range": f"{chr(65 + remarks_col)}{row_index}", "values": [[reason]]},
                {"range": f"{chr(65 + location_col)}{row_index}", "values": [[location]]}
            ])

            # for cell, value in updates:
            #     worksheet.update(range_name=cell, values=value)
            msg = f"✅ Qued update for {name}'s record in {sheet_name} sheet (Row {row_index})"
            print(msg)
            message += f"{msg}\n"

        # Batch update if any
        if updates:
            try:
                worksheet.batch_update(updates)
                print(f"✅ Successfully updated {sheet_name} sheet.")
            except Exception as e:
                success = False
                msg = f"⚠️ Error during batch update in {sheet_name}: {e}"
                print(msg)
                message += f"{msg}\n"

    if success:
        msg = "✅ All updates completed!"
    else:
        msg = "⚠️ Error: Check logs for issue..."
    print(msg)
    message += f"{msg}\n"
    await send_telegram_message(message)
    return success

async def update_informal_sheet(informal_status, names, date_text, informal_sheets_to_update):
    success, message = True, ""

    for sheet_name in informal_sheets_to_update:
        worksheet = informal_sheet.worksheet(sheet_name)
        data = worksheet.get_all_values()
        headers = data[1]  # Use second row as headers
        df = pd.DataFrame(data[2:], columns=headers)  # Data starts from the third row
        print(f"Accessing sheet: '{sheet_name}'")

        # Normalize headers by stripping leading/trailing whitespace
        formatted_headers = [header.strip() for header in headers]

        # Check if the date is a range
        if "-" in date_text:
            # Normalize spaces around "-" and split the range
            start_date, end_date = [d.strip() for d in date_text.split("-")]

            # Determine AM or PM from both start and end dates
            start_am, start_pm = "(AM)" in start_date.upper(), "(PM)" in start_date.upper()
            end_am, end_pm = "(AM)" in end_date.upper(), "(PM)" in end_date.upper()
        else:            
            # Determine AM or PM
            start_am, start_pm = "(AM)" in date_text.upper(), "(PM)" in date_text.upper()
            end_am, end_pm = False, False
        
        # Collect all updates in a batch
        updates = []

        # Update each person's record
        for name in names:
            # Split the name into parts (tokens)
            name_parts = name.split()

            # Attempt to match each part of the name against the Excel sheet
            for part in name_parts:
                print(part)
                matching_rows = df[df["Name"].str.contains(part, case=False, na=False)].index.tolist()

                # Check if there's exactly one match (to confidently identify the person)
                if len(matching_rows) == 1:
                    print(f"✅ Match found: '{name}' matched with row index {matching_rows[0] + 3}")
                    break
                elif len(matching_rows) == 0:
                    print(f"⚠️ No matching name found in {sheet_name} sheet for '{name}'")
                else:
                    # Ambiguous matches (multiple rows match the name tokens)
                    print(f"⚠️ Multiple matches found for '{part}' in {sheet_name} sheet: {matching_rows}")

            if not matching_rows or informal_status == "Invalid":
                success = False
                print(f"⚠️ No matching name found in {sheet_name} sheet for '{name}'" if not matching_rows else f"⚠️ Error: Status {informal_status} for {name} is not valid.")
                continue

            row_index = matching_rows[0] + 3  # Adjusting for header rows
            
            # Extract the days in the date range
            days = extract_days(date_text)
            # Iterate through the days and add updates for each day
            for day in days:
                tdy = datetime(datetime.now().year, datetime.now().month, int(day))
                weekday = tdy.weekday()  # Monday = 0, Sunday = 6

                if weekday == 5 or weekday == 6:  # Saturday or Sunday
                    print(f"📅 Day {day} is a weekend, skipping update.")
                    continue
                elif len(informal_sheets_to_update) == 2 and sheet_name == informal_sheets_to_update[0] and day == days[0] and start_pm:
                    print(f"📅 Status starts at 'PM' for {day}, skipping 'AM'...")
                    continue
                elif len(informal_sheets_to_update) == 2 and sheet_name == informal_sheets_to_update[-1] and day == days[-1] and end_am:
                    print(f"📅 Status ends at 'AM' for {day}, skipping 'PM'...")
                    continue
                # else:
                #     print(f"📅 Day {day} is a weekday, processing...")

                try:
                    date_col = get_column_letter(formatted_headers.index(day)) # Index of day column
                except ValueError:
                    success = False
                    msg = f"⚠️ Error: Column for day '{day}' not found in {sheet_name} sheet."
                    print(msg)
                    message += f"{msg}\n" 
                    continue
                
                # Update the Google Sheet
                updates.extend([
                    {"range": f"{date_col}{row_index}", "values": [[informal_status]]},
                ])
            msg = f"✅ Qued update '{informal_status}' for {name}'s record in {sheet_name} sheet (Row {row_index})"
            print(msg)
            message += f"{msg}\n"
        # Batch update if any
        if updates:
            try:
                worksheet.batch_update(updates)
                print(f"✅ Successfully updated {sheet_name} sheet.")
            except Exception as e:
                success = False
                msg = f"⚠️ Error during batch update in {sheet_name}: {e}"
                print(msg)
                message += f"{msg}\n"

    if success:
        msg = "✅ All updates completed!"
        print(msg)
        message += f"{msg}\n"
    else:
        msg = "⚠️ Error: Check logs for issue..."
        print(msg)
        message += f"{msg}\n"
    await send_telegram_message(message)
    return success

# Step 8: Check for expired status
async def check_and_update_status():
    sheets = ["AM", "PM", "NIGHT"]
    stay_in_ppl = {"Lin Jiarui", "Lee Yang Xuan",
                   "Zhang Haoyuan", "Ong Jun Wei",
                   "Thong Wai Hung", "Lim Jia Hao",
                   "Alfred Leandro Liang", "Haziq Syahmi Bin Norzaim"}
    tomorrow = datetime.now() + timedelta(days=1)
    tmr = tomorrow.strftime("%d/%m/%y")
    weekday = tomorrow.weekday()  # Monday = 0, Sunday = 6
    message = ""
    if weekday == 4:  # Friday
        stay_in_ppl = set()
        print(f"📅 Tomorrow is Friday! Updating all 'STAY IN' statuses to 'P - STAY OUT' for {len(stay_in_ppl)} personel.") # Clear stay-in list so no one stays in
    elif weekday == 5:  # Saturday
        print("📅 Tomorrow is Saturday! No updates needed.")
        return None # Exit function, skipping updates
    elif weekday == 6:  # Sunday
        sheets = ["NIGHT"] # Only update NIGHT sheet
        print("📅 Tomorrow is Sunday! Updating NIGHT sheet only.")
    else:
        print("📅 Tomorrow is a weekday.")
    print(f"Checking statuses for {tmr}...")
    message += f"Checking statuses for {tmr}...\n"

    for sheet_name in sheets:
        print(f"🔎 Accessing worksheet: '{sheet_name}'")
        names, stay_in_names = [], []
        worksheet = sheet.worksheet(sheet_name)
        data = worksheet.get_all_values()
        headers = data[1]  # Use second row as headers
        df = pd.DataFrame(data[2:], columns=headers)  # Data starts from the third row

        # Find the first occurrence of AE
        first_ae_index = df[df["Platoon"] == "AE"].index.min()
        status = "PRESENT" if sheet_name != "NIGHT" else "P - STAY OUT"
        if pd.isna(first_ae_index):  # If no AE platoon members found
            print(f"⚠️ No AE platoon members found in {sheet_name} sheet.")
            continue

        for i, row in df.iloc[first_ae_index:].iterrows():
            platoon, name, date_range, current_status = row["Platoon"], row["Name"], row["Date"].strip(), row["Status"]
            if platoon != "AE": # Stops when no longer AE ppl
                break
            elif not date_range: # Skips ppl with no date
                if name in stay_in_ppl and weekday == 6 and current_status == "P - STAY OUT":
                    stay_in_names.append(name)
                    print(f"🚨 Expired status: {name}, Stay out")
                    message += (f"🚨 Expired status: {sheet_name} | Name: {name} | Status: {row['Status']} | Dates: {row['Date']}\n")
                continue

            # Formate date for comparison
            date_range = date_range.replace("(AM)", "").replace("(PM)", "").strip()
            # print(date_range)
            try:
                # end_date = datetime.strptime(date_range.split("-")[-1].strip(), "%d/%m/%y")
                date_parts = date_range.split("-")
                end_date = datetime.strptime(date_parts[-1].strip(), "%d/%m/%y") if len(date_parts) > 1 else datetime.strptime(date_parts[0].strip(), "%d/%m/%y")
                # print(end_date)

                # Compare end_date to tomorrows's date
                if end_date.date() < tomorrow.date():
                    print(f"🚨 Expired status: {name}")
                    message += (f"🚨 Expired status: {sheet_name} | Name: {name} | Status: {row['Status']} | Dates: {row['Date']}\n")
                    if name in stay_in_ppl:
                        stay_in_names.append(name)
                    else:
                        names.append(name)
            except ValueError: # Skip invalid dates
                print(f"⚠️ Invalid date format for {name}: '{date_range}'")
                continue
        if not message:
            message += f"🔎 No expired status found in {sheet}!\n"
        await send_telegram_message(message)
        message = ""

        # Update each sheet in batches
        # Combine name list for one batch update
        names += stay_in_names
        if names:
            await update_sheet(status, "", names, "", "", [sheet_name])
    # Changes stay out to stay in for those needed
    if stay_in_names:
        await update_sheet("P - STAY IN SGC 377", "", stay_in_names, "", "", ["NIGHT"])

    msg = f"✅ Status check complete! \n📅 Next run scheduled at: {scheduler.get_jobs()[0].next_run_time.strftime('%d/%m/%y %H:%M:%S')}"
    print(msg) # Debugging
    return msg

# Step 9: Run the checks everyday (Cannot be asnyc)
def run_asyncio_task():
    asyncio.run(check_and_update_status())

# Function to start the scheduler
scheduler = BackgroundScheduler(timezone=ZoneInfo("Asia/Singapore")) # Adjust timezone
async def start_scheduler():
    print("Starting scheduler...")
    # scheduler.add_job(lambda: asyncio.create_task(check_and_update_status()), "cron", hour=22, minute=30, misfire_grace_time=60)
    scheduler.add_job(run_asyncio_task, "cron", hour=22, minute=30, misfire_grace_time=60)
    scheduler.start()

    # Ensure job is added before accessing it
    await asyncio.sleep(1)  # Add a small delay to ensure job registration

    jobs = scheduler.get_jobs()
    if jobs and jobs[0].next_run_time:
        next_run_message = f"📅 Next status check will run at: {jobs[0].next_run_time.strftime('%d/%m/%y %H:%M:%S')}"
        print(next_run_message)
        await send_telegram_message(next_run_message)
    else:
        print("⚠️ No scheduled jobs or next run time not available.")