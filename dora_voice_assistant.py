import os

import speech_recognition as sr
import pyttsx3
import pywhatkit as pk
import datetime
import wikipedia as wi
from GoogleNews import GoogleNews
import pyjokes
import webbrowser
import pyautogui
import winshell
import ctypes

class dora:
    name = ''

    def setName(self, name):
        self.name = name


dora_command_list = ["Sorry, I did not get what you said", " I am Dora your assistant, how can I help you?","I did not get that", "Sorry, the service is down"]

def title():
    print("|**********************************************************|")
    print("|Tool     : Dora Voice Assistant                           |")
    print("|Author   : Areed Ahmed Arshad                             |")
    print("|Date     : 11/01/2021                                     |")
    print("|Language : Python based assistant                         |")
    print("|**********************************************************|\n")

def dora_listener():
    listener = sr.Recognizer()
    with sr.Microphone() as source:
        listener.adjust_for_ambient_noise(source)
        print(">> Dora Listening....")
        voice = listener.listen(source)
        try:
            command = listener.recognize_google(voice)
            low_command = command.lower()
            print(">> Your Request: " + low_command)
            if "dora" in low_command:
                return command
            else:
                return 0
        except sr.UnknownValueError:
            dora_speaking(dora_command_list[2])
        except sr.RequestError:
            dora_speaking(dora_command_list[3])


def dora_speaking(command):
    engine = pyttsx3.init()
    engine.setProperty("rate", 150)
    # voices = engine.getProperty('voices') 
  
    # for voice in voices: 
    #     # to get the info. about various voices in our PC  
    #     print("Voice:") 
    #     print("ID: %s" %voice.id) 
    #     print("Name: %s" %voice.name) 
    #     print("Age: %s" %voice.age) 
    #     print("Gender: %s" %voice.gender) 
    #     print("Languages Known: %s" %voice.languages) 
    voice_id = "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Speech\\Voices\\Tokens\\TTS_MS_EN-GB_HAZEL_11.0"
    engine.setProperty("voice", voice_id)
    print(">> Dora Reply: " + command)
    engine.say(command)
    engine.runAndWait()


def dora_responses(user_to_dora_command):
    dora_command_to_give_lower = user_to_dora_command.lower()
    dora_command_to_give = dora_command_to_give_lower.replace('dora', '')

    if there_exists(["play the song", "play song"], dora_command_to_give):
        song_to_play = dora_command_to_give.split("song")[-1]
        dora_speaking("Playing " + song_to_play + " from youtube")
        pk.playonyt(song_to_play)

    elif there_exists(["how are you", "how are you doing"], dora_command_to_give):
        return "I'm very well, thanks for asking " + dora_object.name

    elif there_exists(["what's the time", "tell me the time", "what time is it"], dora_command_to_give):
        what_time = datetime.datetime.now().strftime('%I:%M %p')
        return "Current time is " + what_time + " " + dora_object.name

    elif there_exists(["weather"], dora_command_to_give):
        url = "https://www.google.com/search?sxsrf=ACYBGNSQwMLDByBwdVFIUCbQqya-ET7AAA%3A1578847393212&ei=oUwbXtbXDN-C4-EP-5u82AE&q=weather&oq=weather&gs_l=psy-ab.3..35i39i285i70i256j0i67l4j0i131i67j0i131j0i67l2j0.1630.4591..5475...1.2..2.322.1659.9j5j0j1......0....1..gws-wiz.....10..0i71j35i39j35i362i39._5eSPD47bv8&ved=0ahUKEwiWrJvwwP7mAhVfwTgGHfsNDxsQ4dUDCAs&uact=5"
        webbrowser.get().open(url)
        return "Here is what I found for on google " + dora_object.name

    elif there_exists(["what is my exact location"], dora_command_to_give):
        url = "https://www.google.com/maps/search/Where+am+I+?/"
        webbrowser.get().open(url)
        return "You must be somewhere near here, as per Google maps " + dora_object.name

    elif there_exists(["open gmail"],dora_command_to_give):
        webbrowser.open_new_tab("https://www.gmail.com")
        return "Gmail is open"

    elif there_exists(["price of", "search for"], dora_command_to_give):
        search_term = dora_command_to_give.split("for")[-1]
        url = "https://google.com/search?q=" + search_term
        webbrowser.get().open(url)
        return "Here is what I found for " + search_term + " on google for you " + dora_object.name

    elif there_exists(["capture", "my screen", "screenshot"], dora_command_to_give):
        myScreenshot = pyautogui.screenshot()
        myScreenshot.save('D:/screenshot/dora_screen.png')
        return "Screenshot saved in the screenshot folder of your D drive " + dora_object.name

    elif there_exists(["information on"], dora_command_to_give):
        info = dora_command_to_give.split("on")[-1]
        person = wi.summary(info, 3)
        return "Here is what I found for" + person + "on google " + dora_object.name

    elif there_exists(["news on"], dora_command_to_give):
        new = dora_command_to_give.split("on")[-1]
        dora_news(new)

    elif there_exists(["joke"], dora_command_to_give):
        return pyjokes.get_joke()

    elif there_exists(["open excel", "excel"], dora_command_to_give):
        dora_speaking("Opening microsoft excel for you " + dora_object.name)
        os.startfile('C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE')

    elif there_exists(["open word", "word"], dora_command_to_give):
        dora_speaking("Opening microsoft word for you " + dora_object.name)
        os.startfile('C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE')

    elif there_exists(["open powerpoint", "powerpoint", "power point"], dora_command_to_give):
        dora_speaking("Opening microsoft powerpoint for you " + dora_object.name)
        os.startfile('C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE')

    elif there_exists(["open notepad", "notepad"], dora_command_to_give):
        dora_speaking("Opening notepad for you " + dora_object.name)
        open("dora_notepad.txt", "w")
        os.system("notepad.exe dora_notepad.txt")

    elif there_exists(["shutdown system"], dora_command_to_give):
        dora_speaking("Hold On a Sec ! Your system is on its way to shut down")
        subprocess.call('shutdown / p /f')

    elif there_exists(["empty recycle bin"], dora_command_to_give):
        winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
        dora_speaking("Recycle Bin Recycled")

    elif there_exists(["restart"], dora_command_to_give):
        subprocess.call(["shutdown", "/r"])

    elif there_exists(["hibernate" , "sleep"], dora_command_to_give):
        dora_speaking("Hibernating")
        subprocess.call("shutdown / h")

    elif there_exists(["log off", "sign out"], dora_command_to_give):
        dora_speaking("Make sure all the application are closed before sign-out")
        time.sleep(5)
        subprocess.call(["shutdown", "/l"])

    elif there_exists(["lock window"] , dora_command_to_give):
        dora_speaking("locking the device")
        ctypes.windll.user32.LockWorkStation()

    else:
        return "The mentioned task has not been configured yet" + dora_object.name


def dora_news(news):
    googlenews = GoogleNews()
    googlenews.set_period('5d')
    googlenews.search(news)
    news_result = googlenews.results(sort=True)
    dora_speaking(dora_object.name + ", here are the latest updates from multiple news channels on " + news)
    count = 0
    for r in news_result:
        count = count + 1
        full_description = []
        for k in r.keys():
            if k == "media":
                full_description.append(r[k])
            if k == "desc":
                full_description.append(r[k])
        if full_description[1] != "" and count <= 6:
            final = "News from media " + full_description[0] + " is: " + full_description[1]
            dora_speaking(final)


def there_exists(terms, voice_data):
    for term in terms:
        if term in voice_data:
            return True

# def user_name():
#     bob_speaking(bob_command_list[4])
#     uname = bob_listener()
#     uname_final = uname.replace("bob", "").lower()
#     print(uname_final)
#     bob_speaking("Is your name, " + uname)
#     user_confirmation = bob_listener()
#     user_result = user_confirmation.replace("bob", "").lower()
#     print(user_result)
#     if there_exists(["yes", "true", "sure", "bingo", "absolutely"], user_confirmation):
#         return uname
#     else:
#         user_name()

if __name__ == "__main__":
    try:
        title()
        dora_object = dora()
        dora_object.name = "Master Red"
        owner_name = dora_object.name
        flag = 0
        while True:
            user_to_dora_co = dora_listener()
            if user_to_dora_co == 0:
                dora_speaking(dora_command_list[0])
            else:
                flag = flag + 1
                if flag == 1:
                    if owner_name != "" :
                        dora_object.name = owner_name
                        day_time = int(datetime.datetime.now().strftime('%H'))
                        if day_time < 12:
                            dora_speaking("Hey " + dora_object.name +", Good morning,")
                        elif 12 <= day_time < 18:
                            dora_speaking("Hey " + dora_object.name +", Good Afternoon,")
                        else:
                            dora_speaking("Hey " + dora_object.name +", Good Evening,")
                        dora_speaking(dora_command_list[1])
                else:
                    if there_exists(["exit", "quit", "goodbye", "good bye"], user_to_dora_co.lower()):
                        dora_speaking("going offline " + dora_object.name)
                        exit()
                    else:
                        result = dora_responses(user_to_dora_co)
                        dora_speaking(result)
    except:
        pass
