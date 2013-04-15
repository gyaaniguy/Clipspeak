import pyttsx

import threading
from threading import *
import datetime
import os
import sys
import ctypes
from ctypes import wintypes
import win32clipboard 
import win32con

import time
from Queue import Queue

queue =   Queue()
pauselocation =  [0]
wordsToSay = ''
reduceSpeedBy = 20       
HOTKEYS = {
    1 : (win32con.VK_F3, win32con.MOD_WIN),
    2 : (win32con.VK_F4, win32con.MOD_WIN),
    3 : (win32con.VK_F2, win32con.MOD_WIN)
}

def runNewThread(wordsToSay, startingPoint):    
    global queue, pauselocation
    
    e1 = (queue, wordsToSay, pauselocation, startingPoint)
    t1 = threading.Thread(target=saythread,args=e1)
    #t1.daemon = True 
    t1.start()
    

def handle_win_f3 ():
    
    print 'opening'
    os._exit(1)
    print 'opened'

play = 0 

def handle_win_f4 ():
    global queue
    global pauselocation
    global play, wordsToSay 
    
    if play == 0:
        print 'pausing'
        play = 1
        queue.put(True)
                
    elif play == 1:
        print 'playing'
        play = 0 
        print type(pauselocation) ,  pauselocation
        startingPoint = 0
        for i in pauselocation[1:]:
            startingPoint += i
       # wordsToSay =  wordsToSay[pauselocation[-1]:]
       # print wordsToSay 
        runNewThread(wordsToSay,startingPoint)
        
def handle_win_f2 ():
    global queue
    global pauselocation
    global play, wordsToSay 
   # print 'pausing'
    queue.put(True)
    time.sleep(2)
    #print 'playing'
    #print type(pauselocation) ,  pauselocation
    startingPoint = 0
    for i in pauselocation:
        startingPoint += i
    print 'beofre reducing', startingPoint , wordsToSay[startingPoint:]      
    startingPoint -= 50
   # wordsToSay =  wordsToSay[pauselocation[-1]:]
   # print wordsToSay 
   
    print 'after reducing', startingPoint , wordsToSay[startingPoint:]
    runNewThread(wordsToSay,startingPoint)
        
        
HOTKEY_ACTIONS = {
    1 : handle_win_f3,
    2 : handle_win_f4,
    3 : handle_win_f2
}


def saythread(queue , text , pauselocation, startingPoint):
    global reduceSpeedBy
    #print type(pauselocation) ,  pauselocation[-1]
    #print type(saythread.pauselocation) ,  saythread.pauselocation[0]
    saythread.pauselocation = pauselocation
    saythread.pause = 0 
    #print 'saythread' ,startingPoint
    saythread.engine = pyttsx.init()
      
    saythread.pausequeue1 = False
    
    def onWord(name, location, length):
        #print 'onWord'
        
        
        saythread.pausequeue1  = queue.get(False) 
        #print  'passed to queue- ', saythread.pausequeue1
        saythread.pause = location
        saythread.pauselocation.append(location)
 
        if saythread.pausequeue1 == True :
            saythread.engine.stop()
                
    def onFinishUtterance(name, completed):
        if completed == True:
            os._exit(0)            
                
    def engineRun():
        #print text
         
        if len(saythread.pauselocation) == 1:
            rate = saythread.engine.getProperty('rate')
            print rate 
            saythread.engine.setProperty('rate', rate-reduceSpeedBy)
#    
#        
        textMod = text[startingPoint:]
        #print "startingPoint " , startingPoint , textMod
        
        saythread.engine.say(text[startingPoint:])
        token = saythread.engine.connect("started-word" , onWord )
        saythread.engine.connect("finished-utterance" , onFinishUtterance )
        saythread.engine.startLoop(True)
        #engine.disconnect(token)
    
    engineRun()
    
    #pauselocation =  saythread.pause
    if saythread.pausequeue1 == False:
        #print 'exiting from thread1'
        os._exit(1) 
   # print 'after everything'


    
if __name__ == '__main__':

    wordsToSay = sys.argv[1:]
    
    if len(wordsToSay) == 0:
        win32clipboard.OpenClipboard() 
        try:
            wordsToSay = win32clipboard.GetClipboardData(win32con.CF_TEXT)
            win32clipboard.CloseClipboard()
            if wordsToSay == None:
                wordsToSay = "Copy some text to the clipboard"
        except TypeError as e:
            wordsToSay = "Copy some text to the clipboard"
         
        
    else:
        wordsToSay = " ".join(wordsToSay)       
    print wordsToSay ;
    runNewThread(wordsToSay,0)
    time.sleep(1)
    
    byref = ctypes.byref
    user32 = ctypes.windll.user32
    
        
    for id, (vk, modifiers) in HOTKEYS.items ():
        print "Registering id", id, "for key", vk
        if not user32.RegisterHotKey (None, id, modifiers, vk):
            print "Unable to register id", id
    
    msg = wintypes.MSG()
    while  user32.GetMessageA (byref (msg), None, 0, 0) != 0  :
        if msg.message == win32con.WM_HOTKEY:
            action_to_take = HOTKEY_ACTIONS.get(msg.wParam)
            #print action_to_take
            if action_to_take:
                action_to_take ()
   # words = raw_input("what do you want me to say?")