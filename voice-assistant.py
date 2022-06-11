import subprocess
# connect many process and obtain return outputs
import pyttsx3
# text to speech3
import json
# encode and decode the jason format
import random
# generate pseudo-random numbers with various common distributions
import operator
# fn corresponding to std operators 
import speech_recognition as sr
# used to convert 
import datetime
# to tell current time
import wikipedia
# extract data from wikipedia
import webbrowser
# controller for webbrowser like youtube and google
import os
# miscellaneous os interfaces
import winshell
# for accessing special folders, for using the shell’s file copy, rename & delete functionality, and a certain amount of support for structured storage
import pyjokes
# for extracting jokes from internet
import feedparser
# downloading and parsing syndicated feeds
import smtplib
# smtp protocol client(requires socket)
import ctypes
# It provides C compatible data types, and allows calling functions in DLLs or shared libraries
import time
# helps in defining sleep function
# delay in two consecutive commands
import requests
# for making HTTP requests to a specified URL.
import shutil
# high level file operations using copying
import smtplib
# for sending emails
from twilio.rest import Client
# helps to send message
from clint.textui import progress
# from ecapture import ecapture as ec
from bs4 import BeautifulSoup
# beautifulSoup used for pulling data from HTML and XML files.
import win32com.client as wincl
# API to convert text into speech
from urllib.request import urlopen
# used for handling URL module for python

engine = pyttsx3.init('sapi5')
# to use computer voice provided by microsoft
voices = engine.getProperty('voices')
# print(voices[1].id) 
# 0 for David and 1 for Zira
engine.setProperty('voice', voices[0].id)

def speak(audio):
	engine.say(audio)
	engine.runAndWait()

def wishMe():
	hour = int(datetime.datetime.now().hour)
	if hour>= 0 and hour<12:
		speak("Good Morning Sir !")

	elif hour>= 12 and hour<18:
		speak("Good Afternoon Sir !")

	else:
		speak("Good Evening Sir !")

	assname =("Jarvis 1 point o")
	speak("I am your Assistant")
	speak(assname)
	
def username():
	speak("How can i help you sir")
	uname = takeCommand()
	speak("Welcome Sir")
	speak(uname)
	columns = shutil.get_terminal_size().columns
	speak("How can i Help you Sir")

def takeCommand():
	r = sr.Recognizer()
	with sr.Microphone() as source:
		
		print("Listening...")
		r.pause_threshold = 1
		# 1 sec to fill gap between command

		# r.energy_threshold = 500
		# minimum energy required for recording
		audio = r.listen(source)

	try:
		#  to avoid error while listening voice
		print("Recognizing...")
		query = r.recognize_google(audio, language ='en-in')
		print(f"User said: {query}\n")

	except Exception as e:
		print(e)
		print("Unable to Recognize your voice.")
		return "None"
	
	return query

if __name__ == '__main__':
	clear = lambda: os.system('cls')
	
	# This Function will clean any
	# command before execution of this python file
	clear()
	wishMe()
	username()
	
	while True:
		
		query = takeCommand().lower()
		
		# All the commands said by user will be
		# stored here in 'query' and will be
		# converted to lower case for easily
		# recognition of command

		if 'wikipedia' in query:
    		# wikipedia of Narendra modi
			speak('Searching Wikipedia...')
			query = query.replace("wikipedia", "")
			results = wikipedia.summary(query, sentences = 4)
			speak("According to Wikipedia")
			print(results)
			speak(results)

		elif 'open youtube' in query:
			speak("Here you go to Youtube\n")
			webbrowser.open("youtube.com")

		elif 'open google' in query:
			speak("Here you go to Google\n")
			webbrowser.open("google.com")

		elif 'open stack overflow' in query:
			speak("Here you go to Stack Over flow.Happy coding")
			webbrowser.open("stackoverflow.com")

		elif 'play music' in query or "play song" in query:
			speak("Here you go with music")
			music_dir = "C:\\Users\\lalit agarwal\\Desktop\\B.TECH DOCUMENTS\\B.TECH 3RD SEMESTER\\Mini Project"
			songs = os.listdir(music_dir)
			print(songs)
			random = os.startfile(os.path.join(music_dir, songs[2]))

		elif 'the time' in query:
			
			time=datetime.datetime.now().strftime("%I o clock and %M %p")
           
			speak(time)

		elif 'open chrome' in query:
			codePath = r"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome"
			os.startfile(codePath)

		elif 'how are you' in query:
			speak("I am fine, Thank you")
			speak("How are you, Sir")

		elif 'fine' in query or "good" in query:
			speak("It's good to know that your fine")

		elif "change my name to" in query:
			query = query.replace("change my name to", "")
			assname = query

		elif "change name" in query:
			speak("What would you like to call me, Sir ")
			assname = takeCommand()
			speak("Thanks for naming me")

		elif "what's your name" in query or "What is your name" in query:
			speak("My friends call me")
			assname="jarvis"
			speak(assname)
			print("My friends call me", assname)

		elif 'exit' in query:
			speak("Thanks for giving me your time")
			exit()

		elif "who made you" in query or "who created you" in query:
			speak("I have been created by you sir.")
			
		elif 'joke' in query:
			speak(pyjokes.get_joke())
			
		elif 'search' in query or 'play' in query:
			query = query.replace("search", "")
			query = query.replace("play", "")		
			webbrowser.open(query)

		elif "who i am" in query:
			speak("If you talk then definitely your human.")

		elif "why you came to world" in query:
			speak("Thanks to you sir. further It's a secret")

		elif 'powerpoint presentation' in query:
			speak("opening Power Point presentation")
			power = r"C:\\Users\\lalit agarwal\\Desktop\\B.TECH DOCUMENTS\\B.TECH 3RD SEMESTER\\AI VOICE ASSISTANT.pptx"
			os.startfile(power)

		elif 'is love' in query:
			speak("It is 7th sense that destroy all other senses")

		elif "who are you" in query:
			speak("I am your virtual assistant created by you sir")

		elif 'reason for you' in query:
			speak("I was created as a Mini project by Mister you sir ")

		elif 'change background' in query:
			ctypes.windll.user32.SystemParametersInfoW(20,0,"C:\\Users\\lalit agarwal\\Downloads\\Blue Chill-3.jpg",0)
			speak("Background changed successfully")

		elif 'lock window' in query:
				speak("locking the device")
				ctypes.windll.user32.LockWorkStation()

		elif 'shutdown system' in query:
				speak("Hold On a Sec ! Your system is on its way to shut down")
				subprocess.call('shutdown / p /f')
				
		elif 'empty recycle bin' in query:
			winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
			speak("Recycle Bin Recycled")

		elif "don't listen" in query or "stop listening" in query:
			speak("for how much time you want to stop jarvis from listening commands")
			a = int(takeCommand())
			time.sleep(a)
			print(a)

		elif "where is" in query:
			query = query.replace("where is", "")
			location = query
			speak("User asked to Locate")
			speak(location)
			webbrowser.open("https://www.google.nl / maps / place/" + location + "")

		elif "restart" in query:
			subprocess.call(["shutdown", "/r"])
			
		elif "hibernate" in query or "sleep" in query:
			speak("Hibernating")
			subprocess.call("shutdown / h")

		elif "log off" in query or "sign out" in query:
			speak("Make sure all the application are closed before sign-out")
			time.sleep(5)
			subprocess.call(["shutdown", "/l"])

		elif "write a note" in query:
			speak("What should i write, sir")
			note = takeCommand()
			file = open('jarvis.txt', 'w')
			speak("Sir, Should i include date and time")
			snfm = takeCommand()
			if 'yes' in snfm or 'sure' in snfm:
				strTime = datetime.datetime.now().strftime("% H:% M:% S")
				file.write(strTime)
				file.write(" :- ")
				file.write(note)
			else:
				file.write(note)
		
		elif "display note" in query:
			speak("Showing Notes")
			file = open("jarvis.txt", "r")
			print(file.read())
			speak(file.read(6))

		elif "jarvis" in query:
			
			wishMe()
			speak("Jarvis 1 point o in your service Mister")
			# speak(assname)

		elif "wikipedia" in query:
			webbrowser.open("wikipedia.com")

		elif "Good Morning" in query:
			speak("A warm" +query)
			speak("How are you Mister")
			speak(assname)

		# most asked question from google Assistant
		elif "will you be my gf" in query or "will you be my bf" in query:
			speak("I'm not sure about, may be you should give me some time")

		elif "how are you" in query:
			speak("I'm fine, glad you me that")

		elif "i love you" in query:
			speak("It's hard to understand")