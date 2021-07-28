import speech_recognition as sr
import pyttsx3
import time
import pywhatkit
import os
import datetime
import wikipedia
import pyjokes
from googletrans import Translator
import requests
import smtplib
from gtts import gTTS
import docx2txt
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import json
import subprocess

listener =sr.Recognizer()
engine=pyttsx3.init()
# voices=engine.getProperty('voices')
# engine.setProperty('voice',voices[1].id)
def talk(text):
    engine.say(text)
    engine.runAndWait()
    return text


def take_command():
    try:
        with sr.Microphone() as source:
            print("Listening...")
            voice=listener.listen(source)
            command=listener.recognize_google(voice)
            command =command.lower()
            if 'fox' in command:
                command=command.replace('fox','')
                #print(command)

    except:
        pass
    return command


def weather():

    # Enter your API key here
    talk("Enter the api key")
    api_key=input("Enter the api key")
    # base_url variable to store url
    base_url = "http://api.openweathermap.org/data/2.5/weather?"

    # Give city name
    with sr.Microphone() as source:
        listener.adjust_for_ambient_noise(source, duration=0.2)
        city=talk("Please tell the city")
        ci_name=listener.listen(source)
        city_name=listener.recognize_google(ci_name)
        #city_name = input("Enter city name : ")

    # complete_url variable to store
    # complete url address
    complete_url = base_url + "appid=" + api_key + "&q=" + city_name

    # get method of requests module
    # return response object
    response = requests.get(complete_url)

    # json method of response object
    # convert json format data into
    # python format data
    x = response.json()

    # Now x contains list of nested dictionaries
    # Check the value of "cod" key is equal to
    # "404", means city is found otherwise,
    # city is not found
    if x["cod"] != "404":

        # store the value of "main"
        # key in variable y
        y = x["main"]

        # store the value corresponding
        # to the "temp" key of y
        current_temperature = y["temp"]

        # store the value corresponding
        # to the "pressure" key of y
        current_pressure = y["pressure"]

        # store the value corresponding
        # to the "humidity" key of y
        current_humidity = y["humidity"]

        # store the value of "weather"
        # key in variable z
        z = x["weather"]

        # store the value corresponding
        # to the "description" key at
        # the 0th index of z
        weather_description = z[0]["description"]

        # print following values
        print(" Temperature (in kelvin unit) = " +
              str(current_temperature) +
              "\n atmospheric pressure (in hPa unit) = " +
              str(current_pressure) +
              "\n humidity (in percentage) = " +
              str(current_humidity) +
              "\n description = " +
              str(weather_description))
        talk("Current weather in" +city_name)
        talk("is" +weather_description)

    else:
        print(" City Not Found ")

def transla():
    translator = Translator()
    talk("Please select the source and destination language")
    from_lang=input("Enter the src language")
    to_lang=input("Enter the dst language")
    with sr.Microphone() as source:
        listener.adjust_for_ambient_noise(source, duration=0.2)
        talk("Please tell the phrase")
        audio = listener.listen(source)
        my_text = listener.recognize_google(audio)
        print(my_text)
    text_to_translate = translator.translate(my_text,  src= from_lang,dest= to_lang)
    text = text_to_translate.text
    print(text)
    return text
    #speak.save("captured_voice.mp3")

    # Using OS module to run the translated voice.
    #os.system("start captured_voice.mp3")


def mail():
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    talk("Please enter the mail id ")
    mail_id=input("Enter the mail id")
    talk("Please enter the password")
    mail_pass=input("Enter the password")
    s.login(mail_id, mail_pass)
    talk("Please enter the receiver mail id")
    rev_id=input("Enter the receiver mail id")
    #talk("Please tell the message to be sent")
    with sr.Microphone() as source:
        listener.adjust_for_ambient_noise(source, duration=0.2)
        talk("Please tell the message")
        audio = listener.listen(source)
        message= listener.recognize_google(audio)
        print(message)
    s.sendmail(mail_id, rev_id, message)
    s.quit()

def match_resume():
    talk("Please upload the resume you want to match and the job description")
    res=("C:/Users/hp/OneDrive/Desktop/New_York.docx")
    job_des=("C:/Users/hp/OneDrive/Desktop/jobdes.docx")
    resume=docx2txt.process(res)
    job_description=docx2txt.process(job_des)

    text=[resume,job_description]
    cv=CountVectorizer()
    count_matrix=cv.fit_transform(text)
    print("Similarity Scores:")
    print(cosine_similarity(count_matrix))
    match=cosine_similarity(count_matrix)[0][1]*100
    match=round(match,2)
    print(match)
    talk("Resume matches by")
    talk(match)
    talk("percent")

def application():
    app=talk("Please tell from the below applications you want to open")
    print("\n\t 1.MICROSOFT WORD \t 2.MICROSOFT POWERPOINT \n\t 3.MICROSOFT EXCEL \t 4.GOOGLE CHROME \n\t 5.VLC PLAYER \t 6.ADOBE PDF Viewer\n\t 7.MOZILLA FIREFOX \n\t 8.NOTEPAD \t 9.MICROSOFT EDGE \n\n\t\t")
    command = take_command()
    if "microsoft word" in command:
        talk("Opening Microsoft Word")
        subprocess.Popen("C:/Program Files (x86)/Microsoft Office/Office16/WINWORD.exe")
    elif "microsoft powerpoint" in command:
        talk("Opening Microsoft Powerpoint")
        subprocess.Popen("C:/Program Files (x86)/Microsoft Office/Office16/POWERPNT.exe")
    elif "microsoft excel" in command:
        talk("Opening Microsoft Excel")
        subprocess.Popen("C:/Program Files (x86)/Microsoft Office/Office16/EXCEL.exe")
    elif "google chrome" in command:
        talk("Opening Google Chrome")
        subprocess.Popen("C:/Program Files/Google/Chrome/Application/chrome.exe")
    elif "player" in command:
        talk("Opening VLC Player")
        subprocess.Popen("C:/Program Files (x86)/VideoLAN/VLC/vlc.exe")
    elif "adobe" in command:
        talk("Opening Adobe PDF viewer")
        subprocess.Popen("C:/Program Files (x86)/Adobe/Reader 11.0/Reader/AcroRd32.exe")
    elif "browser" in command:
        talk("Opening Mozilla Firefox")
        subprocess.Popen("C:/Program Files/Mozilla Firefox/firefox.exe")
    elif "notepad" in command:
        talk("Opening Notepad")
        subprocess.Popen("C:/Windows/System32/notepad.exe")
    elif "microsoft edge" in command:
        talk("Opening Microsoft Edge")
        subprocess.Popen("C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe")
    else:
        talk("Sorry Not Able To Recognize")


def run_fox():
    command=take_command()
   # print(command)
    if 'play' in command:
        song=command.replace('play','')
        talk('playing'+song)
        pywhatkit.playonyt(song)
    elif 'time' in command:
        time = datetime.datetime.now().strftime('%I:%M %p')
        print(time)
        talk('Current time is:' + time)

    elif 'search' in command:
        person=command.replace('search','')
        info=wikipedia.summary(person,sentences=1)
        print(info)
        talk(info)
    elif 'tell me a joke' in command:
        #jokes=command.replace('tell me a joke','')
        joke=talk(pyjokes.get_joke())
        print(joke)
    elif 'weather' in command:
        weather()
        #talk("Current weather in %s is %s"+temp)

    elif 'translate' in command:
        tex=transla()
        talk(tex)

    elif 'send mail' in command:
        mail()

    elif 'check resume' in command:
        match_resume()

    elif 'open application' in command:
        application()


talk("Hello there this is Fox,how can I help you")
run_fox()




