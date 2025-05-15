from huggingface_hub import TextToImageTargetSize
from telegram.ext import Application, CommandHandler, MessageHandler, filters, CallbackQueryHandler, ContextTypes
from telegram.constants import ParseMode
from telegram import ChatInviteLink, InlineKeyboardButton, InlineKeyboardMarkup, Update
from contextlib import asynccontextmanager
from google.oauth2.service_account import Credentials
from starlette.requests import ClientDisconnect
from fastapi import FastAPI, Request, Response
from dotenv import load_dotenv
from http import HTTPStatus
import pandas as pd
import tempfile
import calendar
import gspread
import base64
import re
import os
from apscheduler.schedulers.background import BackgroundScheduler
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import asyncio
import time

# Load environment variables from .env
load_dotenv()
# Icons ✅❌⚠️⏭️🔄📩🔔📌🪪📍👥🎭🧑‍🤝‍🧑📅🗓️📝📄📊📋⌛

# Telegram Bot Token
TELEGRAM_TOKEN = os.getenv('Telegram_Token')
CHAT_ID = os.getenv('Chat_ID')
GROUP_CHAT_ID = os.getenv('Group_Chat_ID')
chat_id = GROUP_CHAT_ID # Default
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
# google_sheets_url = os.getenv("google_sheets_url") # AI Sheet
real_google_sheets_url = os.getenv("real_google_sheets_url")
# sheet = client.open_by_url(google_sheets_url) # Change to toggle
sheet = client.open_by_url(real_google_sheets_url) # Change to toggle
# informal_google_sheets_url = os.getenv("informal_google_sheets_url") # AI Sheet
real_informal_google_sheets_url = os.getenv("real_informal_google_sheets_url")
# informal_sheet = client.open_by_url(informal_google_sheets_url) # Change to toggle
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

# Send message when bot starts (Both chats)
async def send_startup_message():
    # Replace with the chat ID where you want to send the message
    chat_id = CHAT_ID  # Can be your own chat ID or a group chat ID
    group_chat_id = GROUP_CHAT_ID
    await ptb.bot.send_message(chat_id, "Startup complete!")
    # await ptb.bot.send_message(group_chat_id, "Startup complete!")
    asyncio.create_task(start_scheduler())
    print("✅ Scheduler started")

# /Start command handler
async def start(update: Update, _: ContextTypes.DEFAULT_TYPE):
    text = '''Haii Haii, you either used this command to test it out or to actually figure out what this bot does.

In short, this bot is meant to help make updating parade state more convenient, yay~
    
How? 
Firstly, you copy the status message from whatsapp to here.
Pls check that the format of the status message is correct because some people just dont like to follow the given format...
    
After you send the message, (Give the bot a few seconds to boot up if its not already active), the bot should reply with the relevant extracted information such as the persons name, status, date, etc.
    
Check, and please check, before pressing the '✅ Confirm' button.
    
The bot will then start updating the relevant excel sheets. You can see what sheet the bot is updating, who and what status is being updated and the excel row and name that the update goes to. (For confirmation that the right row is being updated)

There are regular reminders to tell you to say that strength is updated, so don't ignore it. :/

At the end of the day, 10:30pm, the bot automatically does a check for tomorrows status and clears any expired statuses.

This bot is made in python and extracts information from status messages by matching keywords found. So please, use the format...
    
Hope this helps, for more information on the formatting of status messages, please type '/eg' for exmaples and more details. Ty~'''

    await update.message.reply_text(text=text)
ptb.add_handler(CommandHandler("start", start))

# /id Get chat id
async def get_chat_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Reply with the user's chat ID
    chat_id = update.message.chat.id
    await update.message.reply_text(f"Your chat ID is: {chat_id}")
ptb.add_handler(CommandHandler("id", get_chat_id))

# /check Manually run status check
async def check_status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # telegram_message = await ptb.bot.send_message(chat_id=CHAT_ID, text="🔄 Checking status...")
    telegram_message = await ptb.bot.send_message(chat_id=chat_id, text="🔄 Checking status...")
    message = await check_and_update_status()
    await telegram_message.edit_text(message)

    await asyncio.sleep(2) # Just a pause

    telegram_message = await ptb.bot.send_message(chat_id=chat_id, text="🔄 Checking informal status...")
    message = await check_and_update_informal_status()
    await telegram_message.edit_text(message)
ptb.add_handler(CommandHandler("check", check_status))

# /help Command list
async def command_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = '''📋 List of commands
/help - 🔔 Opens up command list

/start - ✅ Introduction of StatusUpdate Bot

/eg - 🪪 Example formats and explantions (TBC)

/id - 📩 Get chat id of current chatgroup (Debugging purposes)

/check - 🔄 Runs the status check for official and informal excel sheets.
(Before 8pm: Checks if status expired TODAY. After 8pm: Checks if status expires TOMORROW.)'''

    await ptb.bot.send_message(chat_id=chat_id, text=text)
ptb.add_handler(CommandHandler("help", command_list))
# Function for other functions to send Telegram message
async def send_telegram_message(message: str, chat_id: int):
    await ptb.bot.send_message(chat_id=chat_id, text=message)

async def eg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = '''You actually need this? Ok, i'll go through the details with you...
    
📄 *Format*
Status: {Your status} @ {Location you will be at}
R/Name: {Rank and Name}
{Any more Ranks and Names}
Date: {Duration of status}
Location: {Location if any}
Reason: {Reason / Remarks for status if any}

Obviously there are many variations. I'll give examples for those later.

The format above is the general format for most statuses. Anything in {Brackets} is the actual information that you are supposed to send.

*Status*
Start with the word "Status". Spaces and ":" are optional.
Write your status, short form or spelled out correctly pls.
@ If you want to include the location right after.
Eg: "Status: Outstation *@* KC3"
I cannot stress how important this little "@" is. The bot won't detect a location if you don't add this...

*R/Name*
Start with "R/Name". Spaces and ":" are optional again, "R/Name" or "R/Names" is fine too. 
Write your rank and name, yes rank is needed.
Eg: "PTE"/"LCP"/"CPL"/"3SG"/"ME1"
One rank and name per line ok?
Eg: 
"R/Names: CPL Isaac
3SG Isaac Lam"

*Date*
Start with "Date" with or without "s", you get the drill.
There is a specific formats for dates.
Two ways, a string of digits, or with "/" in between.
If the status lasts for multiple days, use a "-" / "to" in between the start and end date.
Eg: "150525" Means 15th May 2025
"15/05/25" --> Say but with "/"
"051525 - 101525" --> 5th to 10th May 2025
"05/15/25 to 10/15/25" --> Same thing and you can use "to" instead
Please *don't* spell it out like "15th May 2025".

*Location*
Aaaaaa im so tired i'll continue this next time, gud night...
'''

    await ptb.bot.send_message(chat_id=chat_id, text=text)
ptb.add_handler(CommandHandler("eg", eg))

@asynccontextmanager
async def lifespan(_: FastAPI):
    await ptb.bot.deleteWebhook()  # Ensure webhook is reset
    await asyncio.sleep(1)  # Small delay to ensure completion
    await ptb.bot.setWebhook("https://updatestatus-production.up.railway.app/webhook") # replace <your-webhook-url>

    # Debugging
    # webhook_info = await ptb.bot.getWebhookInfo()
    # print(f"✅ Webhook set to: {webhook_info.url}")

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
    try:
        # print("✅ Webhook received!")
        req = await asyncio.wait_for(request.json(), timeout=10)
        # print(f"🔎 Raw request data: {req}") # Debugging
    except asyncio.TimeoutError:
        print("⌛ Timeout while recieving request.")
        return Response(status_code=HTTPStatus.REQUEST_TIMEOUT)
    except ClientDisconnect:
        print("⌛ Client disconnected before request was completed.")
        return Response(status_code=HTTPStatus.BAD_REQUEST)
    
    update = Update.de_json(req, ptb.bot)

    global chat_id # Declare global variable
    # Validate if update.message exists
    if update.message is not None:
        chat_id = update.message.chat.id
        print(f"🔔 Chat ID updated: {chat_id}")
    else:
        print("⏭️ Update does not contain a message. Skipping chat_id update.")

    await ptb.process_update(update)
    return Response(status_code=HTTPStatus.OK)

# Message handler for processing status updates
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message.text.strip()
    sender = update.message.from_user.id
    # Check if the message starts with "status" and has sufficient length
    if len(message) < 5 or message[0:6].lower() != "status":
        print("⏭️ Message received is not a status message, skipping...")
        return None
    print(f"📩 From Chat: {chat_id} | User {sender}: \n{message}") # Debugging

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
    
    response = (
        f"✅ *Status Update Received*\n"
        f"📌 *Status:* {status}\n"
        f"🎭 *Informal Status:* {informal_status}\n"
        f"📍 *Location:* {location}\n"
        f"👥 *Names:* {', '.join(names) if names else 'None'}\n"
        f"📅 *Dates:* {date_text}\n"
        f"📝 *Reason:* {reason}\n"
        f"📄 *Sheets to Update:* {', '.join(sheets_to_update)}\n"
        f"📋 *Informal Sheets:* {', '.join(informal_sheets_to_update)}"
    )

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
        print("⚠️ Error: No data found in context.")
        return
    
    # Update excel sheet
    complete = await update_sheet(status, location, names, date_text, reason, sheets_to_update, chat_id)
    success = await update_informal_sheet(informal_status, names, date_text, informal_sheets_to_update, chat_id)
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
    "CCL": "LEAVE",
    "CHILD": "LEAVE",
    "OL": "LEAVE",
    "OVERSEAS": "LEAVE",
    "LEAVE": "LEAVE",
    "TIME OFF": "TO",
    "OFF": "OFF",
    "RSI": "RSI/ RSO",
    "RSO": "RSI/ RSO",
    "MC": "MC",
    "MA": "MA",
    "TO": "TO",
    "C": "CSE",
    "L": "LEAVE",
    "O": "OFF"
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
    "CCL": "CCL",
    "CHILD": "CCL",
    "OL": "OL",
    "OVERSEAS": "OL",
    "LEAVE": "L",
    "TIME OFF": "TO",
    "OFF": "O",
    "RSI": "RSI",
    "RSO": "RSO",
    "TO": "TO",
    "MC": "MC",
    "MA": "MA",
    "XWB": "XWB",
    "XTW": "XTW",
    "L": "L",
    "C": "C",
    "O": "O",
    "PRESENT": 1,
    "1": 1,
    1 : 1
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
    raw_date_text = date_match.group(1).strip() if date_match else ""
    # print(f"Raw date text: {raw_date_text}") # Debugging

    # Convert date to fixed format (DD/MM/YY)
    # Regular expression for 6-digit (DDMMYY) dates
    six_digit_pattern = r"\b(\d{1,2})[\/]?(\d{1,2})[\/]?(\d{2})\b"

    # Check if it's a range (date-date or date - date)
    sheets_to_update, informal_sheets_to_update = [], []
    informal_sheet_name = datetime.now().strftime("%b %y")

    # Format range dates
    if "to" in raw_date_text:
        raw_date_text = re.sub(r"\s*to\s*", " - ", raw_date_text)
    am_pattern, pm_pattern = r"\b(AM)\b", r"\b(PM)\b"

    if "-" in raw_date_text:
        # Normalize spaces around "-" and split the range
        start_date, end_date = [d.strip() for d in raw_date_text.split("-")]

        # Determine AM or PM from both start and end dates
        start_period, end_period = "", ""
        if re.search(am_pattern, start_date.upper(), re.IGNORECASE):
            start_period = " (AM)"
        elif re.search(pm_pattern, start_date.upper(), re.IGNORECASE):
            start_period = " (PM)"
        if re.search(am_pattern, end_date.upper(), re.IGNORECASE):
            end_period = " (AM)"
        elif re.search(pm_pattern, end_date.upper(), re.IGNORECASE):
            end_period = " (PM)"

        # Format dates
        date_match = re.search(six_digit_pattern, start_date)
        if date_match:
            start_date = f"{date_match.group(1)}/{date_match.group(2)}/{date_match.group(3)}"
            # print(start_date)
        else:
            print("No valid date found.")
        date_match = re.search(six_digit_pattern, end_date)
        if date_match:
            end_date = f"{date_match.group(1)}/{date_match.group(2)}/{date_match.group(3)}"
            # print(end_date)
        else:
            print("No valid date found.")

        # Convert individual dates
        start_date = re.sub(six_digit_pattern, r"\1/\2/\3", start_date)
        end_date = re.sub(six_digit_pattern, r"\1/\2/\3", end_date)
        date_text = f"{start_date}{start_period} - {end_date}{end_period}"

        # date_text is a range, all sheets need to be updated
        sheets_to_update.extend(["AM", "PM"])
        informal_sheets_to_update.extend([f"{informal_sheet_name} (AM)", f"{informal_sheet_name} (PM)"])        
    else:
        # Format single date
        date_match = re.search(six_digit_pattern, raw_date_text)
        if date_match:
            date_text = f"{date_match.group(1)}/{date_match.group(2)}/{date_match.group(3)}"
            # print(date_text)
        else:
            print("No valid date found.")

        # Determine AM or PM
        # Combine text inputs to check for AM or PM
        combined_text = f"{raw_status} {location} {raw_date_text}".upper()

        # Determine AM or PM
        start_am, start_pm = False, False
        if re.search(am_pattern, combined_text):
            start_am = True
            date_text += " (AM)"
        elif re.search(pm_pattern, combined_text):
            start_pm = True
            date_text += " (PM)"
        
        # Convert date
        date_text = re.sub(six_digit_pattern, r"\1/\2/\3", date_text)

        # Determine sheets to update
        if start_am:
            sheets_to_update.append("AM")
            informal_sheets_to_update.append(f"{informal_sheet_name} (AM)")
        if start_pm:
            sheets_to_update.append("PM")
            informal_sheets_to_update.append(f"{informal_sheet_name} (PM)")
        if len(sheets_to_update) == 0:
            sheets_to_update.extend(["AM", "PM"])
            informal_sheets_to_update.extend([f"{informal_sheet_name} (AM)", f"{informal_sheet_name} (PM)"])

    # Night sheet updates only for specific statuses
    if status in ["DUTY", "CSE", "AO", "LEAVE", "OFF", "MC"] and (len(sheets_to_update) == 2 or sheets_to_update[0] == "PM"):
        sheets_to_update.append("NIGHT")

    # Extract Location and Reason (if provided separately)
    reason = "" # Location already has a check
    for line in lines:
        location_match = re.match(r"Locations?\s*:?\s*(.*)", line, re.IGNORECASE)
        if location_match:
            location = location_match.group(1).strip()
        
        remark_match = re.match(r"Remarks?\s*:?\s*(.*)", line, re.IGNORECASE)
        reason_match = re.match(r"Reasons?\s*:?\s*(.*)", line, re.IGNORECASE)
        if reason_match:
            reason = reason_match.group(1).strip()
        elif remark_match:
            reason = remark_match.group(1).strip()

        mc_no_match = re.match(r"MC No.?\s*:?\s*(.*)", line, re.IGNORECASE)
        if mc_no_match:
            reason = mc_no_match.group(1).strip()

    # Output extracted values
    # print("Extracted Raw Status:", raw_status)
    print("Extracted Status:", status)
    print("Extracted Informal Status:", informal_status)
    print("Extracted Names:", names)
    # print("Extracted Raw Date:", raw_date_text)
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
    day_pattern = r'(\d{1,2})/(\d{2})/(\d{2})'
    current_year, current_month = int(datetime.now().year % 100), int(datetime.now().month)

    # Search for all matches of the day pattern
    date = re.findall(day_pattern, date_text)
    # print(f"Date: {date}\nDay: {date[0][0]}")

    if len(date) == 1: # If there is only one day, return that day
        return [str(int(date[0][0]))]
    elif len(date) == 2: # If there are two dates, generate the range of days
        # Extract days, months, and years
        start_day, start_month, start_year = map(int, date[0])
        end_day, end_month, end_year = map(int, date[1])

        # Check month and year
        if start_month != current_month or start_year != current_year:
            start_day = 1
        elif end_month != current_month or end_year != current_year:
            end_day = calendar.monthrange(current_year, current_month)[1]
        
        # Generate all days within the range
        day_list = []
        current_day = start_day
        while current_day <= end_day:
            day_list.append(str(current_day)) # Append the day as a string
            current_day += 1
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

def find_name_index(df, name, sheet_name, official):
    if official:
        matching_rows = df[(df["Name"].str.contains(name, case=False, na=False)) & (df["Platoon"] == "AE")].index.tolist()
    else:
        matching_rows = df[df["Name"].str.contains(name, case=False, na=False)].index.tolist()
    
    # 1. Direct Substring Match
    if len(matching_rows) == 1:
        row_index = matching_rows[0] + 3  # Adjusting for header rows
        print(f"✅ Direct match found: '{name}' matched in row {row_index} | {df['Name'][matching_rows[0]]}")
        return row_index
    
    # Squishing names by removing spaces
    df["Squished_Name"] = df["Name"].str.replace(" ", "")    

    # 2. Name Part Match
    name_parts = name.split()
    part_matches = []  # Store all rows that match name parts
    for part in name_parts:
        # print(part)
        if official:
            partial_matching_rows = df[(df["Squished_Name"].str.contains(part, case=False, na=False)) & (df["Platoon"] == "AE")].index.tolist()
        else:
            partial_matching_rows = df[df["Squished_Name"].str.contains(part, case=False, na=False)].index.tolist()
        part_matches.extend(partial_matching_rows)

        if len(partial_matching_rows) == 1:  # Found an exact match for a part
            row_index = partial_matching_rows[0] + 3  # Adjusting for header rows
            print(f"✅ Part match found: '{part}' matched in row {row_index} | {df['Name'][partial_matching_rows[0]]}")
            return row_index
        elif len(partial_matching_rows) == 0:
            print(f"⚠️ No matching name found for '{part}' in '{sheet_name}' sheet.")
        else:
            print(f"⚠️ Multiple matches found for '{part}' in '{sheet_name}' sheet: {partial_matching_rows}.")

    # 3. Most Common Row Saved
    if part_matches:
        # Count occurrences of each row index
        row_count = {}
        for index in part_matches:
            row_count[index] = row_count.get(index, 0) + 1

        # Select the row index with the highest count
        max_count = max(row_count.values())
        most_common_rows = [index for index, count in row_count.items() if count == max_count]

        # Check for ties
        if len(most_common_rows) > 1:
            print(f"⚠️ Equal matches found for '{name}' in rows: {most_common_rows}. Skipping...")
        else:
            row_index = most_common_rows[0] + 3  # Adjusting for header rows
            print(f"✅ Most common match found for '{name}' in row {row_index} | {df['Name'][most_common_rows[0]]}")
            return row_index

    # No matches found
    print(f"⚠️ No valid matches found for '{name}' in '{sheet_name}' sheet.")
    return None

# Step 7: Update Google Sheets for each sheet
async def update_sheet(status, location, names, date_text, reason, sheets_to_update, chat_id):
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
            platoon_col = formatted_headers.index("Platoon")
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
            row_index = find_name_index(df, name, sheet_name, official=True)
            if row_index == None:
                message += f"⚠️ No valid matches found for '{name}' in '{sheet_name}' sheet.\n"
                continue

            # Update the Google Sheet
            updates.extend([
                {"range": f"{chr(65 + status_col)}{row_index}", "values": [[status]]},
                {"range": f"{chr(65 + date_col)}{row_index}", "values": [[date_text]]},
                {"range": f"{chr(65 + remarks_col)}{row_index}", "values": [[reason]]},
                {"range": f"{chr(65 + location_col)}{row_index}", "values": [[location]]}
            ])

            # for cell, value in updates:
            #     worksheet.update(range_name=cell, values=value)
            msg = f"⌛ '{status}' -> {name} | {sheet_name} sheet | Name: {df['Name'][row_index-3]}"
            print(msg)
            message += f"{msg}\n"

        # Batch update if any
        if updates:
            try:
                worksheet.batch_update(updates)
                print(f"✅ Successfully updated {sheet_name} sheet.")
            except Exception as e:
                success = False
                msg = f"⚠️ Error during batch update in {sheet_name}: {e}\n⚠️ Retrying update..."
                print(msg)
                message += f"{msg}\n"

    if success:
        msg = "✅ All updates completed!"
    else:
        msg = "⚠️ Error: Check logs for issue..."
    print(msg)
    message += f"{msg}\n"
    await send_telegram_message(message, chat_id=chat_id)
    return success

async def update_informal_sheet(informal_status, names, date_text, informal_sheets_to_update, chat_id):
    success, message = True, ""

    for sheet_name in informal_sheets_to_update:
        worksheet = informal_sheet.worksheet(sheet_name)
        data = worksheet.get_all_values()
        headers = data[1]  # Use second row as headers
        df = pd.DataFrame(data[2:], columns=headers)  # Data starts from the third row
        print(f"🔎 Accessing sheet: '{sheet_name}'")

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
            row_index = find_name_index(df, name, sheet_name, official=False)
            if row_index == None:
                message += f"⚠️ No valid matches found for '{name}' in '{sheet_name}' sheet.\n"
                continue
            
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
                else:
                    print(f"📅 Day {day} is a weekday, processing...")

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
            msg = f"⌛ '{informal_status}' -> {name} | {sheet_name} sheet | Name: {df['Name'][row_index-3]}"
            print(msg)
            message += f"{msg}\n"

        # Batch update if any
        if updates:
            try:
                worksheet.batch_update(updates)
                print(f"✅ Successfully updated '{sheet_name}' sheet.")
            except Exception as e:
                success = False
                msg = f"⚠️ Error during batch update in '{sheet_name}': {e}\n⚠️ Retrying update..."
                print(msg)
                message += f"{msg}\n"
                await asyncio.sleep(2)
                worksheet.batch_update(updates) # Retry once

    if success:
        msg = "✅ All updates completed!"
        print(msg)
        message += f"{msg}\n"
    else:
        msg = "⚠️ Error: Check logs for issue..."
        print(msg)
        message += f"{msg}\n"
    await send_telegram_message(message, chat_id=chat_id)
    return success

# Step 8: Check for expired status
async def check_and_update_status():
    sheets = ["AM", "PM", "NIGHT"]
    stay_in_ppl = {"Lin Jiarui", "Lee Yang Xuan",
                   "Zhang Haoyuan", "Ong Jun Wei",
                   "Thong Wai Hung", "Lim Jia Hao",
                   "Alfred Leandro Liang", "Haziq Syahmi Bin Norzaim"}
    timezone = datetime.now(ZoneInfo("UTC"))
    hour = timezone.astimezone(ZoneInfo("Asia/Singapore")).hour
    tomorrow = timezone
    if hour >= 20:
        tomorrow += timedelta(days=1)
    tmr = tomorrow.strftime("%d/%m/%y")
    weekday = tomorrow.weekday()  # Monday = 0, Sunday = 6
    message = ""
    if weekday == 4:  # Friday
        stay_in_ppl = set()
        print(f"📅 Friday! Updating all 'STAY IN' statuses to 'P - STAY OUT' for {len(stay_in_ppl)} personel.") # Clear stay-in list so no one stays in
    elif weekday == 5:  # Saturday
        print("📅 Saturday! No updates needed.")
        return None # Exit function, skipping updates
    elif weekday == 6:  # Sunday
        sheets = ["NIGHT"] # Only update NIGHT sheet
        print("📅 Sunday! Updating NIGHT sheet only.\nUpdating all stay in personel's status to 'P - STAY IN'.")
    else:
        print("📅 A weekday.")
    print(f"Checking statuses for {tmr}...")
    message += f"Checking statuses for {tmr}...\n"

    for sheet_name in sheets:
        msg = f"🔎 Accessing worksheet: '{sheet_name}'"
        print(msg)
        message += f"{msg}\n"
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
                # Friday stay in to stay out
                if weekday == 4 and current_status == "P - STAY IN SGC 377":
                    names.append(name)
                # Stay outs to stay in (and weekday == 6) removed
                elif name in stay_in_ppl and current_status == "P - STAY OUT":
                    stay_in_names.append(name)
                else:
                    continue
                print(f"🚨 Expired status: {name}")
                message += (f"🚨 Expired status: {sheet_name} | Name: {name} | Status: {row['Status']} | Dates: {row['Date']}\n")
                continue

            # Formate date for comparison
            date_range = date_range.replace("(AM)", "").replace("(PM)", "").strip()
            # print(date_range)
            try:
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
            message += f"🔎 No expired status found in '{sheet_name}'!\n"
        await send_telegram_message(message, chat_id=chat_id)
        message = ""

        # Update each sheet in batches
        # Combine name list for one batch update
        names += stay_in_names
        if names:
            await update_sheet(status, "", names, "", "", [sheet_name], chat_id)
    # Changes stay out to stay in for those needed
    if stay_in_names:
        await update_sheet("P - STAY IN SGC 377", "", stay_in_names, "", "", ["NIGHT"], chat_id)

    msg = f"📅 Next run scheduled at: {scheduler.get_jobs()[0].next_run_time.strftime('%d/%m/%y %H:%M:%S')}"
    print(msg) # Debugging
    await send_telegram_message(msg, chat_id=chat_id)
    return "✅ Status check complete!"

async def check_and_update_informal_status():
    timezone = datetime.now(ZoneInfo("UTC"))
    hour = timezone.astimezone(ZoneInfo("Asia/Singapore")).hour
    tomorrow = timezone
    if hour >= 20:
        tomorrow += timedelta(days=1)
    day = str(int(tomorrow.strftime("%d"))) # Convert "01" to "1" etc
    tmr = tomorrow.strftime("%d/%m/%y")
    weekday = tomorrow.weekday()  # Monday = 0, Sunday = 6
    informal_sheet_name = tomorrow.strftime("%b %y")
    informal_sheets = [f"{informal_sheet_name} (AM)", f"{informal_sheet_name} (PM)"]
    message = ""
    if weekday == 5 or weekday == 6:  
        print("📅 Its a weekend! No updates needed.")
        return None # Exit function, skipping updates
    else:
        print("📅 A weekday.")
    print(f"Checking statuses for {tmr}...")
    message += f"Checking statuses for {tmr}...\n"

    for sheet_name in informal_sheets:
        msg = f"🔎 Accessing worksheet: '{sheet_name}'"
        print(msg)
        message += f"{msg}\n"
        names, default_status = [], 1
        worksheet = informal_sheet.worksheet(sheet_name)
        data = worksheet.get_all_values()
        headers = data[1]  # Use second row as headers
        df = pd.DataFrame(data[2:], columns=headers)  # Data starts from the third row

        # Get the indexes of S/N columns
        second_name_batch = df[df["S/N"] == "S/N"].index.min()
        # print(f"✅ Second occurrence of 'S/N' is at index {second_name_batch}")
        if second_name_batch is None or pd.isna(second_name_batch):
            print(f"⚠️ No batch found in {sheet_name} sheet.")
            continue
        
        for i, row in df.iloc[second_name_batch:].iterrows():
            sn, name = row["S/N"].strip(), row["Name"].strip()
            # print(f"S/N: {sn} | Name: {name}") # Debugging
            # Skip if sn is not a digit
            if not sn.isdigit():
                continue
            # Check if current status is empty
            if not row[day].strip():
                names.append(name)
                msg = f"🚨 Empty status: {name}"
                print(msg)
                message += f"{msg}\n"
        
        await send_telegram_message(message, chat_id=chat_id)
        message = ""
        
        # Update each sheet in batches
        if names:
            print(f"📝 Updating '{sheet_name}' for names: {names} with status: {default_status}")
            await update_informal_sheet(default_status, names, tmr, [sheet_name], chat_id)
    msg = f"📅 Next run scheduled at: {scheduler.get_jobs()[0].next_run_time.strftime('%d/%m/%y %H:%M:%S')}"
    print(msg) # Debugging
    await send_telegram_message(msg, chat_id=chat_id)
    return "✅ Status check complete!"
        
async def send_reminder():
    global chat_id
    timezone = datetime.now(ZoneInfo("UTC"))
    time = timezone.astimezone(ZoneInfo("Asia/Singapore"))
    day = time.weekday()
    hour = time.hour

    # Check if its a weekend
    if day == 5 or day == 6:
        print("Its a weekend, skipping reminder.")
        return None

    # Determine the period of the day
    if 0 <= hour < 11: # From 12am to before 11am
        period = "AM"
    elif 11 <= hour < 17: # From 11am to before 5pm
        period = "PM"
    elif hour >= 17: # From 5pm onwards
        period = "NIGHT"
    else:
        period = "UNSPECIFIED"  # For times outside the defined periods

    if not chat_id:
        print("⚠️ No valid chat_id found, skipping reminder.")
        return None

    # Send the reminder
    print("🔔 Sending Reminder...")
    await send_telegram_message(f"🔔 Reminder to update {period} status on WhatsApp~", chat_id)
    return None

# Step 9: Run the checks everyday (Cannot be asnyc)
def run_asyncio_task():
    try:
        loop = asyncio.get_running_loop()
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)

    loop.run_until_complete(check_and_update_status())
    loop.run_until_complete(check_and_update_informal_status())

def run_timed_reminders():
    try:
        loop = asyncio.get_running_loop()
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)

    loop.run_until_complete(send_reminder())

# Function to start the scheduler
scheduler = BackgroundScheduler(timezone=ZoneInfo("Asia/Singapore")) # Adjust timezone
async def start_scheduler():
    print("Starting scheduler...")
    # End of day check
    scheduler.add_job(run_asyncio_task, "cron", hour=22, minute=30, misfire_grace_time=60, coalesce=True, max_instances=1)
    # Update Whatsapp Reminders
    scheduler.add_job(run_timed_reminders, "cron", hour=8, misfire_grace_time=60, coalesce=True, max_instances=1)
    scheduler.add_job(run_timed_reminders, "cron", hour=12, misfire_grace_time=60, coalesce=True, max_instances=1)
    scheduler.add_job(run_timed_reminders, "cron", hour=18, misfire_grace_time=60, coalesce=True, max_instances=1)
    scheduler.start()

    # Ensure job is added before accessing it
    time.sleep(1)  # Add a small delay to ensure job registration

    jobs = scheduler.get_jobs()
    if jobs and jobs[0].next_run_time:
        next_run_time = jobs[0].next_run_time if jobs[0].next_run_time else None
        if next_run_time:
            next_run_message = f"📅 Next status check will run at: {jobs[0].next_run_time.strftime('%d/%m/%y %H:%M:%S')}"
            print(next_run_message)
            # await send_telegram_message(next_run_message, chat_id=CHAT_ID)
            await send_telegram_message(next_run_message, chat_id=CHAT_ID)
        else:
            print("⚠️ No scheduled jobs or next run time not available.")