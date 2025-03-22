import keyboard
from telegram.ext import Application, CommandHandler, MessageHandler, filters, CallbackQueryHandler, ContextTypes
from telegram.constants import ParseMode
from contextlib import asynccontextmanager
from http import HTTPStatus
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext._contexttypes import ContextTypes
from fastapi import FastAPI, Request, Response
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
import gspread
import pandas as pd
import base64
import tempfile
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
google_sheets_url = os.getenv("google_sheets_url")
sheet = client.open_by_url(google_sheets_url)
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

# /Start command handler
async def start(update: Update, _: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Starting...")
ptb.add_handler(CommandHandler("start", start))

# /id Get chat id
async def get_chat_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Reply with the user's chat ID"""
    chat_id = update.message.chat_id
    await update.message.reply_text(f"Your chat ID is: {chat_id}")
ptb.add_handler(CommandHandler("id", get_chat_id))

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
    status, location, names, date_text, reason, sheets_to_update = extract_message(message)

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
                    f"📍 Location: {location}\n"
                    f"👥 Names: {', '.join(names) if names else 'None'}\n"
                    f"📅 Dates: {date_text}\n"
                    f"📄 Reason: {reason}\n")
    await update.message.reply_text(response, reply_markup=reply_markup, parse_mode="Markdown"), 

    # Wait for user to confirm update
    context.user_data["status_data"] = (status, location, names, date_text, reason, sheets_to_update)
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
    await query.message.reply_text("✅ Confirmed! Status updated.")

    # Get data and update sheet
    status, location, names, date_text, reason, sheets_to_update = context.user_data.pop('status_data', (None,)*6)

    if not status:
        print("Error: No data found in context.")
        return
    
    # Update excel sheet
    complete = update_sheet(status, location, names, date_text, reason, sheets_to_update)
    if complete:
        print("✅ All updates completed!")
    else:
        print("⚠️ Error: Check logs for issue...")
ptb.add_handler(CallbackQueryHandler(handle_confirmation))

# Step 5: Define Official Status Mapping
status_mapping = {
    "PRESENT": "PRESENT",
    "ATTACH IN": "ATTACH IN",
    "DUTY": "DUTY",
    "UDO": "DUTY",
    "CDOS": "DUTY",
    "GUARD": "DUTY",
    "WFH": "WFH",
    "OUTSTATION": "OUTSTATION",
    "BLOOD DONATION": "OUTSTATION",
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
    status_match = re.search(r"Status\s*:?\s*([A-Z]+(?:\s+[A-Z]+)?)\s*(?:\s*@\s*(.+))?$", message, re.IGNORECASE | re.MULTILINE)
    raw_status = status_match.group(1).strip() if status_match else "Unknown"
    location = status_match.group(2).strip() if status_match and status_match.group(2) else ""

    # Convert status to official version
    for keyword in status_mapping.keys():
        if keyword.upper() in raw_status.upper():  # Case-insensitive matching
            status = status_mapping[keyword]  # Return mapped status
            break
        else:
            status = "Invalid"

    if status == "Invalid":
        print(f"⚠️ Error: '{raw_status}' is not a valid status.")

    # Extract Names (Handles "R/Name" with or without ":" and same-line names)
    name_lines = []
    name_section = False

    for line in message.split("\n"):
        line = line.strip()
        
        # If "R/Name" is found, start capturing names
        match = re.match(r"R/Names?\s*:?\s*(.*)", line, re.IGNORECASE)
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

    date_match = re.search(r"Dates?\s*:?\s*(.+)", message, re.IGNORECASE)
    date_text = date_match.group(1).strip() if date_match else ""
    period = ""

    # Convert date to fixed format (DD/MM/YY)
    # Regular expression for 6-digit (DDMMYY) dates
    six_digit_pattern = r"\b(\d{2})(\d{2})(\d{2})\b"

    # Check if it's a range (date-date or date - date)
    if "-" in date_text:
        # Normalize spaces around "-" and split the range
        start_date, end_date = [d.strip() for d in date_text.split("-")]

        # Convert individual dates
        start_date = re.sub(six_digit_pattern, r"\1/\2/\3", start_date)
        end_date = re.sub(six_digit_pattern, r"\1/\2/\3", end_date)

        date_text = f"{start_date} - {end_date}"
    else:
        # Convert single date
        date_text = re.sub(six_digit_pattern, r"\1/\2/\3", date_text)

    if "(AM)" in date_text.upper():
        period = "AM"
    elif "(PM)" in date_text.upper():
        period = "PM"

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
    reason_match = re.search(r"Reasons?\s*:?\s*(.+)", message, re.IGNORECASE)
    reason = reason_match.group(1).strip() if reason_match else ""

    # Extract Location (if provided separately)
    location_match = re.search(r"Locations?\s*:?\s*(.+)", message, re.IGNORECASE)
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
    success = True

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
            success = False
            print(f"⚠️ Error: Required columns missing in {sheet_name} sheet.")
            continue

        # Update each person's record
        for name in names:
            matching_rows = df[df["Name"].str.contains(name, case=False, na=False)].index.tolist()

            if not matching_rows or status == "Invalid":
                success = False
                error_msg = f"⚠️ No matching name found in {sheet_name} sheet for '{name}'" if not matching_rows else f"⚠️ Error: Status {status} for {name} is not valid."
                print(error_msg)
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

    for sheet_name in sheets:
        print(f"🔎 Accessing worksheet: '{sheet_name}'")
        names, stay_in_updates = [], []
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
            platoon, name, date_range = row["Platoon"], row["Name"], row["Date"].strip()
            if platoon != "AE": # Stops when no longer AE ppl
                break
            elif not date_range: # Skips ppl with no date
                continue
            elif sheet_name == "NIGHT" and name in stay_in_ppl:
                stay_in_updates.append(name)
                continue # Separate list for stay ins
            # print(f"📌 {sheet_name} | Row {i+3} | Status: {row['Status']} | Dates: {row['Date']}")

            # Formate date for comparison
            date_range = date_range.replace("(AM)", "").replace("(PM)", "").strip()
            # print(date_range)
            try:
                end_date = datetime.strptime(date_range.split("-")[-1].strip(), "%d/%m/%y")
                # print(end_date)

                # Compare end_date to tomorrows's date
                if end_date < tomorrow:
                    print(f"🚨 Expired status: {name}")
                    names.append(name)
            except ValueError: # Skip invalid dates
                print(f"⚠️ Invalid date format for {name}: '{date_range}'")
                continue

        # Update each sheet in batches
        # print(names, len(names))
        if names:
            update_sheet(status, "", names, "", "", [sheet_name])
    if stay_in_updates:
        update_sheet("P - STAY IN SGC 377", "", stay_in_updates, "", "", ["NIGHT"])

    print(f"✅ Status check complete! \n📅 Next run scheduled at: {scheduler.get_jobs()[0].next_run_time}")

# Step 9: Run the checks everyday
# Function to start the scheduler
scheduler = BackgroundScheduler(timezone=ZoneInfo("Asia/Singapore")) # Adjust timezone
async def start_scheduler():
    print("Starting scheduler...")
    # scheduler = BackgroundScheduler(timezone=ZoneInfo("Asia/Singapore")) # Adjust timezone

    scheduler.add_job(lambda: asyncio.create_task(check_and_update_status()), "cron", hour=22, minute=30, misfire_grace_time=60)
    # scheduler.add_job(lambda: asyncio.run(check_and_update_status()), "cron", hour=22, minute=30, misfire_grace_time=60)
    print(f"📅 Run scheduled at: {scheduler.get_jobs()[0].next_run_time}")
    scheduler.start()
    print("Scheduler is running. Press Ctrl+C to exit.")
    try:
        while True:
            await asyncio.sleep(3600)
    except KeyboardInterrupt:
        print("Shutting down scheduler.")
        scheduler.shutdown()

# Call the scheduler when the app starts
async def main():
    await start_scheduler()

# Run the main function
if __name__ == "__main__":
    asyncio.run(main())  # This ensures the event loop is started and executed