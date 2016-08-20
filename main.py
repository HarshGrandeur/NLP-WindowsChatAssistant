import Tkinter
import os
top = Tkinter.Tk()
from chatterbot import ChatBot
import pyttsx
engine = pyttsx.init()
rate = engine.getProperty('rate')
engine.setProperty('rate', rate-30)
import wiki_bot
chatbot = ChatBot('Chatty')
from chatterbot.training.trainers import ListTrainer
import speech_recognition as sr
r = sr.Recognizer()
m = sr.Microphone()
#set threhold level
with m as source: r.adjust_for_ambient_noise(source)
print("Set minimum energy threshold to {}".format(r.energy_threshold))
import wikipedia


from win32com.client import Dispatch







name="Chatty"
conver = [
    "Hello",
    "Hi there! I am "+name,
    "How are you doing?",
    "I'm doing great.",
    "Where do you live?",
    "I live in Mumbai",
    "Where are you from?",
    "I live in Navi Mumbai",
    "Where in navi mumbai?",
    "Cbd bealpur",
    "Are you robot or human?",
    "I am human ofcourse.",
    "That is good to hear",
    "Me too",
    "Thank you.",
    "You're welcome.",
    
    
]



chatbot.set_trainer(ListTrainer)
chatbot.train(conver)
i=1

while(True):
    with sr.Microphone() as source:
        print("Say something!")
        audio = r.listen(source)
    try:    
        msg=r.recognize_google(audio)
        print msg
    except sr.UnknownValueError:
        print("Google Speech Recognition could not understand audio")
        continue
    except sr.RequestError as e:
        print("Could not request results from Google Speech Recognition service; {0}".format(e))
        continue
    if("what is" in msg):
        
        c=msg[8:]
        #print c
        #x=wiki_bot.get_info(c)
        x=wikipedia.summary(c,sentences=5)
        x=str(x)
        print type(x)
        engine.say(x)
        engine.runAndWait()
        continue
    
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
        cur.execute("SELECT * FROM Jobs")
        row=cur.fetchall()
        for r in row:
            print r
   
    else:
        print con
        engine.say(con)
        engine.runAndWait()
    
