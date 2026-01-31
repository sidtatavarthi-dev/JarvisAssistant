# Import required libraries
import speech_recognition as sr  # For speech recognition
import pyttsx3  # For text-to-speech
import webbrowser  # For opening websites
import random  # For random responses
import time  # For delays
import screen_brightness_control as sbc  # For brightness control
import pyautogui  # For keyboard control and tab closing
import os  # For system operations

# Initialize text-to-speech engine
engine = pyttsx3.init('sapi5')  # Force Windows SAPI5 for better reliability

# Configure voice settings - LOUDER
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)  # Male voice (use voices[1] for female)
engine.setProperty('rate', 165)  # Speed of speech
engine.setProperty('volume', 1.0)  # MAX VOLUME

# Initialize speech recognizer
recognizer = sr.Recognizer()

# Dictionary of commands and their URLs/actions
COMMANDS = {
    "youtube": "https://www.youtube.com",
    "d2l": "https://d2l.brightspace.com",
    "google docs": "https://docs.google.com",
    "google drive": "https://drive.google.com",
    "gmail": "https://mail.google.com",
    "google classroom": "https://classroom.google.com",
    "instagram": "https://www.instagram.com",
    "snapchat": "https://www.snapchat.com",
    "twitter": "https://www.twitter.com",
    "tiktok": "https://www.tiktok.com",
    "netflix": "https://www.netflix.com",
    "spotify": "https://www.spotify.com",
    "discord": "https://www.discord.com",
    "reddit": "https://www.reddit.com",
    "github": "https://www.github.com",
    "linkedin": "https://www.linkedin.com",
    "amazon": "https://www.amazon.com",
    "facebook": "https://www.facebook.com",
    "twitch": "https://www.twitch.com",
    "google": "https://www.google.com",
    "chatgpt": "https://chat.openai.com",
    "stack overflow": "https://stackoverflow.com",
    "wikipedia": "https://www.wikipedia.org",
    "whatsapp": "https://web.whatsapp.com",
}

# Fun response variations
ACKNOWLEDGMENTS = [
    "Got it Sid, coming right up!",
    "On it Sid!",
    "You got it boss!",
    "Opening that for you now!",
    "Right away Sid!",
    "Consider it done!",
    "No problem, opening it now!",
    "Sure thing Sid!",
    "Absolutely, one moment!",
    "Opening that right up!"
]

VOLUME_RESPONSES = [
    "Adjusting volume for you Sid!",
    "Volume control activated!",
    "You got it!",
    "On it boss!",
    "Done!",
]

BRIGHTNESS_RESPONSES = [
    "Adjusting brightness for you Sid!",
    "Brightness control activated!",
    "You got it!",
    "On it boss!",
    "Done!",
]

TAB_RESPONSES = [
    "Closing that tab for you Sid!",
    "Consider it done!",
    "On it!",
    "Tab closed boss!",
]

ERROR_MESSAGES = [
    "Sorry Sid, I didn't catch that. Try again?",
    "Could you repeat that Sid?",
    "Didn't quite get that, say it again?",
    "Sorry, what was that?",
    "Can you say that one more time?"
]

NOT_FOUND_MESSAGES = [
    "Sorry Sid, I don't know how to open that yet.",
    "Hmm, I'm not sure about that one.",
    "That's not in my database yet, Sid.",
    "I don't have that command programmed yet."
]

def speak(text):
    # Make Jarvis speak LOUDLY - with better error handling
    print(f"🗣️ Jarvis: {text}")
    try:
        engine.say(text)
        engine.runAndWait()
    except Exception as e:
        print(f"Speech error: {e}")
        # Try alternate method
        try:
            import win32com.client
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            speaker.Speak(text)
        except:
            print("Could not speak - speech engine error")

def set_volume_level(level):
    # Set system volume to exact level (0-100)
    try:
        # Speak acknowledgment FIRST
        response = random.choice(VOLUME_RESPONSES)
        print(f"🗣️ Jarvis: {response}")
        engine.say(response)
        engine.runAndWait()
        
        # Step 1: Set volume to 0 (mute, then press volume down many times)
        pyautogui.press('volumemute')
        time.sleep(0.2)
        
        # Press volume down 50 times to ensure we're at 0
        for i in range(50):
            pyautogui.press('volumedown')
            time.sleep(0.01)
        
        # Unmute
        pyautogui.press('volumemute')
        time.sleep(0.2)
        
        # Step 2: Now press volume up to reach desired level
        # Each press is approximately 2% volume
        presses = int(level / 2)
        print(f"Pressing volume up {presses} times to reach {level}%")
        
        for i in range(presses):
            pyautogui.press('volumeup')
            time.sleep(0.03)  # Slightly longer delay for accuracy
        
        confirmation = f"Volume set to {level} percent"
        print(f"🗣️ Jarvis: {confirmation}")
        engine.say(confirmation)
        engine.runAndWait()
                
    except Exception as e:
        speak("Sorry, couldn't adjust the volume")
        print(f"Error: {e}")

def set_brightness_level(level):
    # Set screen brightness (0-100)
    try:
        response = random.choice(BRIGHTNESS_RESPONSES)
        print(f"🗣️ Jarvis: {response}")
        engine.say(response)
        engine.runAndWait()
        
        sbc.set_brightness(level)
        
        confirmation = f"Brightness now at {level} percent"
        print(f"🗣️ Jarvis: {confirmation}")
        engine.say(confirmation)
        engine.runAndWait()
    except Exception as e:
        speak("Sorry, couldn't adjust the brightness")
        print(f"Error: {e}")

def close_tab(tab_number):
    # Close a specific browser tab
    try:
        response = random.choice(TAB_RESPONSES)
        print(f"🗣️ Jarvis: {response}")
        engine.say(response)
        engine.runAndWait()
        
        # Use Ctrl+Tab to navigate to the tab, then Ctrl+W to close
        # First go to tab 1
        pyautogui.hotkey('ctrl', '1')
        time.sleep(0.1)
        
        # Navigate to the specific tab
        if tab_number <= 8:
            pyautogui.hotkey('ctrl', str(tab_number))
            time.sleep(0.2)
            # Close the tab
            pyautogui.hotkey('ctrl', 'w')
            
            confirmation = f"Tab {tab_number} closed"
            print(f"🗣️ Jarvis: {confirmation}")
            engine.say(confirmation)
            engine.runAndWait()
        else:
            speak("I can only close tabs 1 through 8")
            
    except Exception as e:
        speak("Sorry, couldn't close that tab")
        print(f"Error: {e}")

def extract_number(text):
    # Extract number from text
    words_to_numbers = {
        "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
        "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
        "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14, "fifteen": 15,
        "sixteen": 16, "seventeen": 17, "eighteen": 18, "nineteen": 19,
        "twenty": 20, "twenty five": 25, "thirty": 30, "forty": 40, "fifty": 50,
        "sixty": 60, "seventy": 70, "eighty": 80, "ninety": 90,
        "hundred": 100
    }
    
    # Try to find a digit number first
    words = text.split()
    for word in words:
        if word.isdigit():
            return int(word)
    
    # Try word numbers
    for word, num in words_to_numbers.items():
        if word in text:
            return num
    
    return None

def listen():
    # Listen for voice input
    with sr.Microphone() as source:
        print("\n🎤 Listening...")
        
        # Adjust for ambient noise
        recognizer.adjust_for_ambient_noise(source, duration=0.5)
        
        try:
            # Listen with timeout
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=5)
            print("🔄 Processing...")
            
            # Convert speech to text using Google Speech Recognition
            text = recognizer.recognize_google(audio)
            print(f"You said: {text}")
            return text.lower()
            
        except sr.WaitTimeoutError:
            return ""
        except sr.UnknownValueError:
            speak(random.choice(ERROR_MESSAGES))
            return ""
        except sr.RequestError:
            speak("Sorry Sid, my speech service is down.")
            return ""

def process_command(command):
    # Process the voice command
    # Check if command starts with "jarvis"
    if "jarvis" not in command:
        return
    
    # Remove "jarvis" from command
    command = command.replace("jarvis", "").strip()
    
    # Check for exit commands
    if any(word in command for word in ["exit", "quit", "bye", "stop", "shutdown"]):
        speak("Shutting down. See you later Sid!")
        return "exit"
    
    # Check for close tab command
    if "close" in command and "tab" in command:
        tab_num = extract_number(command)
        if tab_num is not None and 1 <= tab_num <= 8:
            close_tab(tab_num)
            return
        else:
            speak("Which tab number? Say a number from 1 to 8")
            return
    
    # Check for volume commands
    if "volume" in command:
        if "mute" in command:
            response = random.choice(VOLUME_RESPONSES)
            print(f"🗣️ Jarvis: {response}")
            engine.say(response)
            engine.runAndWait()
            
            pyautogui.press('volumemute')
            speak("Volume muted")
            return
        elif "max" in command or "full" in command or "hundred" in command:
            set_volume_level(100)
            return
        elif "up" in command or "increase" in command or "raise" in command:
            response = random.choice(VOLUME_RESPONSES)
            print(f"🗣️ Jarvis: {response}")
            engine.say(response)
            engine.runAndWait()
            
            for _ in range(5):
                pyautogui.press('volumeup')
                time.sleep(0.05)
            
            speak("Volume increased!")
            return
        elif "down" in command or "decrease" in command or "lower" in command:
            response = random.choice(VOLUME_RESPONSES)
            print(f"🗣️ Jarvis: {response}")
            engine.say(response)
            engine.runAndWait()
            
            for _ in range(5):
                pyautogui.press('volumedown')
                time.sleep(0.05)
            
            speak("Volume decreased!")
            return
        else:
            # Try to extract specific number
            level = extract_number(command)
            if level is not None and 0 <= level <= 100:
                print(f"Setting volume to: {level}%")
                set_volume_level(level)
                return
            else:
                speak("What volume level? Say a number from 0 to 100")
                return
    
    # Check for brightness commands
    if "brightness" in command:
        if "max" in command or "full" in command or "hundred" in command:
            set_brightness_level(100)
            return
        elif "min" in command or "zero" in command or "dark" in command:
            set_brightness_level(10)
            return
        elif "up" in command or "increase" in command or "raise" in command:
            try:
                response = random.choice(BRIGHTNESS_RESPONSES)
                print(f"🗣️ Jarvis: {response}")
                engine.say(response)
                engine.runAndWait()
                
                current_brightness = sbc.get_brightness()[0]
                new_brightness = min(100, current_brightness + 10)
                sbc.set_brightness(new_brightness)
                
                speak(f"Brightness now at {new_brightness} percent")
            except:
                speak("Couldn't get current brightness")
            return
        elif "down" in command or "decrease" in command or "lower" in command:
            try:
                response = random.choice(BRIGHTNESS_RESPONSES)
                print(f"🗣️ Jarvis: {response}")
                engine.say(response)
                engine.runAndWait()
                
                current_brightness = sbc.get_brightness()[0]
                new_brightness = max(10, current_brightness - 10)
                sbc.set_brightness(new_brightness)
                
                speak(f"Brightness now at {new_brightness} percent")
            except:
                speak("Couldn't get current brightness")
            return
        else:
            # Try to extract specific number
            level = extract_number(command)
            if level is not None and 0 <= level <= 100:
                set_brightness_level(level)
                return
            else:
                speak("What brightness level? Say a number from 0 to 100")
                return
    
    # Remove "open", "go to", etc. for website commands
    command = command.replace("open", "").strip()
    command = command.replace("go to", "").strip()
    command = command.replace("launch", "").strip()
    command = command.replace("start", "").strip()
    
    # Look for matching command
    found = False
    for app_name, url in COMMANDS.items():
        if app_name in command:
            # Acknowledge FIRST with forced output
            response = random.choice(ACKNOWLEDGMENTS)
            print(f"🗣️ Jarvis: {response}")
            engine.say(response)
            engine.runAndWait()
            
            # Then open the website
            webbrowser.open(url)
            found = True
            break
    
    if not found and command.strip():
        speak(random.choice(NOT_FOUND_MESSAGES))

def main():
    # Main program loop
    print("=" * 60)
    print("🤖 JARVIS VOICE ASSISTANT")
    print("=" * 60)
    print("\nCommands:")
    print("  - Say: 'Jarvis open [app name]'")
    print("  - Say: 'Jarvis set volume to [0-100]'")
    print("  - Say: 'Jarvis volume up/down'")
    print("  - Say: 'Jarvis set brightness to [0-100]'")
    print("  - Say: 'Jarvis brightness up/down'")
    print("  - Say: 'Jarvis close tab [1-8]'")
    print("  - Say: 'Jarvis exit' to quit")
    print("\nSupported apps:")
    for app in list(COMMANDS.keys())[:10]:
        print(f"  • {app.title()}")
    print(f"  ... and {len(COMMANDS) - 10} more!")
    print("\n" + "=" * 60)
    
    # Initial greeting - LOUD AND CLEAR
    speak("Hello Sid! Jarvis is online and ready.")
    speak("I can open apps, control volume and brightness, and close tabs.")
    
    # Main listening loop
    while True:
        command = listen()
        
        if command:
            result = process_command(command)
            if result == "exit":
                break
        
        # Small delay between commands
        time.sleep(0.5)
    
    print("\n✅ Jarvis shut down successfully.")

# Run the assistant
if __name__ == "__main__":
    main()
