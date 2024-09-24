import os
import random
import re
import difflib

from typing import Final
from telegram import Update, InputFile, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, filters, ContextTypes
import openpyxl

# Your bot token and username
#TOKEN: Final = '7867149104:AAEHVuUah67WSzhDd24VPTK6ou0aq-xHvFM'
#BOT_USERNAME: Final = '@IdiomsUp_bot'
TOKEN: Final = '6991746723:AAEGi-DzARSPgm0F2IJ-y8wKzxp_4PhtmLc'
BOT_USERNAME: Final = '@Aradhya0404_Bot'
IDIOMS_FILE = 'idioms.txt'
EXCEL_FILE = 'user_scores.xlsx'
IDIOMS_EXCEL_FILE = 'idioms_data.xlsx'  # Path to the Excel file containing idiom data

# Dictionary to keep track of ongoing idiom games
idiom_game_state = {}

# Helper to escape MarkdownV2 characters
def escape_markdown_v2(text: str) -> str:
    return re.sub(r'([_\*\[\]\(\)~`>#+\-=|{}.!])', r'\\\1', text)

# Command to show all user scores
async def show_all_results(update: Update, context: ContextTypes.DEFAULT_TYPE):
    scores = load_scores()  # Load scores from the Excel file

    if not scores:
        await update.message.reply_text("No scores found")
        return

    # Build the message to display all users
    message = "*All Users and Scores:*\n"
    for user_id, username, score in scores:
        message += f"ID: {user_id}, Username: {username}, Score: {score} points\n"

    await update.message.reply_text(message, parse_mode='MarkdownV2')

def update_user_score(user_id: int, username: str, score: int):
    # If the file does not exist, create it with headers
    if not os.path.exists(EXCEL_FILE):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = 'Scores'
        # Create the header row
        sheet.append(['Idnumber', 'Username', 'Score'])
        workbook.save(EXCEL_FILE)

    # Load the existing Excel file
    workbook = openpyxl.load_workbook(EXCEL_FILE)
    sheet = workbook.active

    # Check if the user already exists by checking the user ID
    user_found = False
    for row in range(2, sheet.max_row + 1):  # Start from row 2 to skip the header
        if sheet.cell(row=row, column=1).value == user_id:  # The user ID is in the first column (Idnumber)
            # Update the score for this user
            current_score = sheet.cell(row=row, column=3).value
            new_score = current_score + score if current_score is not None else score
            sheet.cell(row=row, column=3, value=new_score)  # Update the score in the 3rd column
            user_found = True
            break

    if not user_found:
        # If user not found, add a new row with the user ID, username, and score
        sheet.append([user_id, username, score])

    # Save the updated Excel file
    workbook.save(EXCEL_FILE)
    workbook.close()

def load_scores():
    if not os.path.exists(EXCEL_FILE):
        return []

    workbook = openpyxl.load_workbook(EXCEL_FILE)
    sheet = workbook.active

    scores = []
    for row in range(2, sheet.max_row + 1):  # Start from row 2 to skip the header
        user_id = sheet.cell(row=row, column=1).value
        username = sheet.cell(row=row, column=2).value
        score = sheet.cell(row=row, column=3).value

        if user_id and username and score is not None:
            scores.append((user_id, username, score))

    workbook.close()
    return scores

# Command to show the top 10 idiom users
async def select_top_10_idiom_users(update: Update, context: ContextTypes.DEFAULT_TYPE):
    scores = load_scores()

    if not scores:
        await update.message.reply_text("No scores found")
        return

    # Sort by score in descending order
    scores.sort(key=lambda x: x[2], reverse=True)

    # Get the top 10 users
    top_10 = scores[:10]

    # Build the message to display top users
    message = "*Top 10 Idiom Users:*\n"
    for idx, (user_id, username, score) in enumerate(top_10, 1):
        message += f"{idx}: {username} : {score} points\n"

    await update.message.reply_text(message, parse_mode='MarkdownV2')

# Command to show the user's rank and score
async def my_rank_in_idiom(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    username = update.message.from_user.username or update.message.from_user.first_name

    scores = load_scores()

    if not scores:
        await update.message.reply_text("No scores found")
        return

    # Sort by score in descending order
    scores.sort(key=lambda x: x[2], reverse=True)

    # Find user's rank
    user_rank = None
    for rank, (u_id, u_name, score) in enumerate(scores, 1):
        if u_id == user_id:
            user_rank = (rank, score)
            break

    if user_rank:
        rank, score = user_rank
        await update.message.reply_text(f"Your rank: {rank}\nYour score: {score}")
    else:
        await update.message.reply_text("You haven't played the idiom game yet")
# Function to get a random idiom from the file
# Function to get a random idiom from the Excel file, avoiding repeats
def get_random_idiom_from_excel(file_path: str, used_srno: list):
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active

        # Collect idioms, meanings, examples, and srno from the Excel file
        idioms_data = []
        for row in range(2, sheet.max_row + 1):  # Start from the second row to skip headers
            srno = sheet.cell(row=row, column=1).value  # 'srno' is in the first column
            if srno in used_srno:  # Skip idioms that have already been used
                continue
            idiom = sheet.cell(row=row, column=2).value  # Assuming idiom is in column 2
            meaning = sheet.cell(row=row, column=3).value  # Assuming meaning is in column 3
            example = sheet.cell(row=row, column=4).value  # Assuming example is in column 4
            image_name = sheet.cell(row=row, column=5).value  # Assuming image name is in column 5

            # Append tuple of srno, idiom, meaning, example, image name
            idioms_data.append((srno, idiom, meaning, example, image_name))

        # Choose a random idiom from the list of unused idioms
        if idioms_data:
            srno, idiom, meaning, example, img = random.choice(idioms_data)
            img_file_path = f'Image/{img.strip()}.jpg' if img else None
            return srno, idiom, meaning, example, img_file_path
        else:
            return None, None, None, None, None

    except FileNotFoundError:
        return None, None, None, None, None

# Start the idiom game and ask how many idioms
async def start_idiom_game_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.message.chat.id

    if chat_id in idiom_game_state:
        await update.message.reply_text("An idiom game is already running in this group.")
        return
    keyboard = [
        [InlineKeyboardButton("5 Idioms", callback_data='5')],
        [InlineKeyboardButton("10 Idioms", callback_data='10')],
        [InlineKeyboardButton("15 Idioms", callback_data='15')],
        [InlineKeyboardButton("20 Idioms", callback_data='20')],
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text('How many idioms do you want?', reply_markup=reply_markup)

def is_similar_idiom_in_message(user_text: str, idiom: str, threshold: float = 0.7) -> bool:
    """
    Check if the user's text contains the idiom with a similarity above the given threshold
    and also ensure the user provides at least two other words besides the idiom.
    """
    user_words = user_text.lower().split()

    # Split the idiom into words
    idiom_words = idiom.lower().split()

    # Ensure the user's message contains at least two extra words besides the idiom
    if len(user_words) <= len(idiom_words) + 1:
        return False  # Not enough extra words

    idiom_length = len(idiom_words)

    # Compare each substring of the user's text with the idiom
    for i in range(len(user_words) - idiom_length + 1):
        substring = ' '.join(user_words[i:i + idiom_length])

        # Calculate the similarity ratio between the substring and the idiom
        similarity = difflib.SequenceMatcher(None, substring, idiom.lower()).ratio()

        # Return True if the similarity is above the threshold
        if similarity >= threshold:
            return True

    # If no similar substring is found, return False
    return False

# Handle the user's choice of idioms count
# Handle the user's choice of idioms count
async def button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    chat_id = query.message.chat.id

    if chat_id in idiom_game_state:
        await query.message.reply_text("An idiom game is already running in this group.")
        return
    
    number_of_idioms = int(query.data)
    chat_id = query.message.chat.id

    # Initialize game state for the group (chat) if not already done
    if chat_id not in idiom_game_state:
        idiom_game_state[chat_id] = {
            'players': {},
            'idioms_left': number_of_idioms,
            'current_idiom': None,
            'used_srno': []  # Initialize used_srno as an empty list to avoid KeyError
        }

    # Send the first idiom
    await send_next_idiom(query.message, chat_id)

# Send the next idiom to the group
# Send the next idiom to the group, ensuring no repeats and checking for image existence
async def send_next_idiom(message, chat_id):
    game = idiom_game_state[chat_id]

    if game['idioms_left'] > 0:
        used_srno = game.get('used_srno', [])  # Get the list of used 'srno'

        srno, idiom, meaning, example, img_file_path = get_random_idiom_from_excel(IDIOMS_EXCEL_FILE, used_srno)

        if idiom:
            game['current_idiom'] = idiom  # Save current idiom
            game['used_srno'].append(srno)  # Add the idiom's srno to the used list

            caption = f"*Idiom*: {idiom}\n>*Meaning*: {meaning}\n*Example*: {example}\n\nMake a sentence using this idiom"

            if img_file_path and os.path.isfile(img_file_path):  # Check if the image file exists
                await message.reply_photo(photo=open(img_file_path, 'rb'), caption=caption, parse_mode='MarkdownV2')
            else:
                # If the image does not exist, just send the text (idiom, meaning, and example)
                await message.reply_text(caption, parse_mode='MarkdownV2')
        else:
            await message.reply_text("Error fetching idiom or image")
    else:
        # Game finished, show results
        await show_game_results(message, chat_id)

# Show the results of the game, including all participants
# Show the results of the game, sorted by score in descending order
async def show_game_results(message, chat_id):
    game = idiom_game_state[chat_id]

    # Sort players by score in descending order and filter out players with a score of 0
    sorted_players = [(user_id, player) for user_id, player in sorted(game['players'].items(), key=lambda x: x[1]['score'], reverse=True) if player['score'] > 0]

    if sorted_players:
        results = "*Game Over Here are the results:*\n"
        for user_id, player in sorted_players:
            results += f"@{player['username']} :: {player['score']} points\n"

        await message.reply_text(results, parse_mode='MarkdownV2')
    else:
        await message.reply_text("No participants scored in this game.")

    # Clear the game state
    del idiom_game_state[chat_id]


# Handle the user's message and check if it contains the idiom
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.message.chat.id
    user_id = update.message.from_user.id
    user_name = update.message.from_user.first_name
    username = update.message.from_user.username or user_name  # Use username if available, fallback to first name

    # Escape username for MarkdownV2
    username = escape_markdown_v2(username)

    if chat_id in idiom_game_state and idiom_game_state[chat_id]['current_idiom']:
        game = idiom_game_state[chat_id]
        current_idiom = game['current_idiom']

        # Add user to the players list if they haven't participated yet
        if user_id not in game['players']:
            game['players'][user_id] = {
                'username': username,  # Save username
                'score': 0
            }

        # Check if the user's message contains a part that matches the idiom with at least 70% similarity
        if is_similar_idiom_in_message(update.message.text, current_idiom):
            game['idioms_left'] -= 1
            game['current_idiom'] = None  # Reset current idiom

            # Update player's score
            game['players'][user_id]['score'] += 1

            # Store the username, user ID, and updated score in the Excel file
            update_user_score(user_id, username, game['players'][user_id]['score'])

            # Send the next idiom or show results if game is finished
            await send_next_idiom(update.message, chat_id)
        else:
            # If the message isn't similar enough, don't send a prompt
            pass
    else:
        # Default handling for messages that aren't part of the idiom game
        return None

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text('Hello! Thanks For Chatting With Me, I am YourBot.')

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text('No worries, I will assist you with all kinds of help. For more help, contact @YourContactUsername.')

async def error(update: Update, context: ContextTypes.DEFAULT_TYPE):
    print(f'Update {update} caused error {context.error}')


if __name__ == "__main__":
    print('Starting bot...')
    app = Application.builder().token(TOKEN).build()

    app.add_handler(CommandHandler('start', start_command))
    app.add_handler(CommandHandler('help', help_command))
    app.add_handler(CommandHandler('topplayer', select_top_10_idiom_users))
    app.add_handler(CommandHandler('myrank', my_rank_in_idiom))
    app.add_handler(CommandHandler('showallresult', show_all_results))
    app.add_handler(CommandHandler('startidiom', start_idiom_game_command))
    app.add_handler(CallbackQueryHandler(button_callback))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_error_handler(error)

    print('Polling the bot...')
    app.run_polling(poll_interval=1)



