import pandas as pd
import smtplib
import random
import os
from datetime import datetime
from email.message import EmailMessage
from gtts import gTTS
from playsound import playsound

# File paths
file_path = r"D:\\BirthdayReminder_AIML\\birthdays.xlsx"
birthday_image_path = r"D:\\BirthdayReminder_AIML\\Wish2.jpeg"
song_folder = r"D:\\BirthdayReminder_AIML\\Songs"

# Email Credentials
EMAIL_SENDER = "your mail id"
EMAIL_PASSWORD = "mail id app password"  # Use an App Password

# Birthday wishes
birthday_wishes = [
    "Hey {name}, it's YOUR day! Shine bright, laugh loud, and make unforgettable memories. HAPPY BIRTHDAY!",
    "Happy Birthday, {name}! May today be as awesome as you are. Time to celebrate BIG!",
    "Boom! Another amazing year for {name}! Wishing you joy, success, and endless fun ahead!",
    "Happy Birthday, {name}! Hope your day is filled with surprises, fun, and all the cake you can eat!",
    "Cheers to another fantastic year, {name}! May your day be as incredible as you are. Have a blast!"
]

try:
    # Read all sheets from the Excel file
    dfs = pd.read_excel(file_path, sheet_name=None)  # Load all sheets as dictionary
    df_list = []

    for sheet_name, df in dfs.items():
        if {'Name', 'Birthday', 'Email', 'Branch'}.issubset(df.columns):
            df['Birthday'] = pd.to_datetime(df['Birthday'], errors='coerce')
            today = datetime.today().strftime("%m-%d")  # Only month-day format
            df['Formatted_Birthday'] = df['Birthday'].dt.strftime("%m-%d")  # Ignore year
            df_list.append(df)

    # Combine all sheets into one DataFrame
    df = pd.concat(df_list, ignore_index=True)
    all_contacts = df[['Name', 'Email']].dropna()
    birthday_people = df[df['Formatted_Birthday'] == today]

    if not birthday_people.empty:
        songs = [os.path.join(song_folder, f) for f in os.listdir(song_folder) if f.endswith('.mp3')]
        
        for _, row in birthday_people.iterrows():
            name = row['Name']
            branch = row['Branch']
            song = random.choice(songs) if songs else None
            wish = random.choice(birthday_wishes).format(name=name)
            
            if song:
                print(f"🎶 Playing random song before announcing {name}'s birthday...")
                playsound(song)
            
            print(f"Playing Birthday wish audio for {name}...")
            tts = gTTS(text=wish, lang='en', tld='co.in')
            tts_file = f"{name}_birthday.mp3"
            tts.save(tts_file)
            playsound(tts_file)
            os.remove(tts_file)

        birthday_names = ', '.join(birthday_people['Name'])
        birthday_branches = ', '.join(birthday_people['Branch'])
        wish_message = (
            f"🎉🎂 Hip, hip, hooray! Today is a very special day because it's {birthday_names} from {birthday_branches} birthday! 🎈🥳\n\n"
            f"Let's all come together to celebrate and shower them with love, laughter, and joy! 🎁✨\n"
            f"{birthday_names}, may your day be filled with happiness, your heart with warmth, and your year ahead with endless opportunities! 💖🌟\n\n"
            f"Drop your best wishes and send lots of virtual hugs their way! 💌🎊 #HappyBirthday{birthday_names}"
        )

        def send_email(to_email, recipient_name):
            try:
                msg = EmailMessage()
                msg['Subject'] = "🎉Trial of Our Python Automated Project (Birthday Alert! Let's Celebrate!) 🎂"
                msg['From'] = EMAIL_SENDER
                msg['To'] = to_email
                msg.set_content(f"\n{wish_message}\n\nBest Regards,\nYour Birthday Reminder Bot")
                
                with open(birthday_image_path, 'rb') as img:
                    msg.add_attachment(img.read(), maintype='image', subtype='jpeg', filename='birthday_wish.jpg')
                
                with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
                    server.login(EMAIL_SENDER, EMAIL_PASSWORD)
                    server.send_message(msg)
                print(f"Email sent at {to_email}")
            except Exception as e:
                print(f"Email failed: {e}")

        for _, row in all_contacts.iterrows():
            send_email(row['Email'], row['Name'])
    else:
        print("No birthdays today.")

except Exception as e:
    print(f"Error: {e}")