import Tkinter  # python interface to the Tk GUI toolkit
import os       # operating system interface
top = Tkinter.Tk()

import pyttsx     # cross-platform text-to-speech
# This code is used to choose the specific synthesizer by name
engine = pyttsx.init() 
rate = engine.getProperty('rate') # get the current value of the engine property(rate - integer speeach rate per minute)
engine.setProperty('rate', rate-30)
import wiki_bot

# speech recognition via microphone
import speech_recognition as sr # 
r = sr.Recognizer()
m = sr.Microphone()
# set threshold level
with m as source : r.adjust_for_ambient_noise(source)
print("Set minimum energy threshold to {}".format(r.energy_threshold))

import wikipedia # python liberary access and parse data from Wikipedia.

# python liberary for windows extention
from win32com.client import Dispatch

while(True):
     with sr.Microphone() as source: # Audio input from user
          print("Say something!")
          audio = r.listen(source)
     try:    
          msg=r.recognize_google(audio) # Speech to string output by google speech recognition API
          print msg
     except sr.UnknownValueError:
          print("Google Speech Recognition could not understand audio")
          continue
     except sr.RequestError as e:
          print("Could not request results from Google Speech Recognition service; {0}".format(e))
          continue
     if("what is" in msg): # If it starts with what is... then print the first 5 lines from wiki for the respective subject
          c=msg[8:]
          #print c
          #x=wiki_bot.get_info(c)
          x=wikipedia.summary(c,sentences=5)
          x=str(x)
          print type(x)
          engine.say(x)
          engine.runAndWait()
          continue
     # Other manual commands for the chat bot
     elif("open" in msg):
          print "Opening file"
          if("Notepad" in msg):
               os.system("start notepad++.exe C:\Program Files\Notepad++")
          elif ("command" in msg):
               os.system("start")
          elif("studio" in msg):
               os.system("start rstudio.exe C:\\Program Files\\RStudio\\bin")
          elif("Firefox" in msg):
               os.system("start firefox.exe")
          elif("arduino" in msg):
               os.system("start  C:/\"Program Files\"/Arduino/arduino.exe")
          elif("Team" in msg):
               os.system("start  C:/\"Program Files\"/TeamViewer/TeamViewer.exe")
          elif("sublime" in msg):
               os.system("start  C:/\"Program Files\"/\"Sublime Text 3\"/sublime_text.exe")
          elif("excel" in msg):    
               xl=Dispatch('Excel.Application')
               wb = xl.Workbooks.Open('C:\\Users\\HRSHB\\Desktop\\crawl.csv')
               xl.visible=True
     response=chatbot.get_response(msg)
     con=str(response)
     temp=con[:5]
     if(temp=="print"):
          exec con
     elif(con[:3]=="cur"):
          print con
          engine.say(con)
          engine.runAndWait()
          row=cur.fetchall()
          for r in row:
               print r   
     else:
          print con
          engine.say(con)
          engine.runAndWait()