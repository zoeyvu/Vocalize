from win32api import *
from win32gui import *
import win32con
import sys, os
import struct
import time
import pyautogui
import speech_recognition as sr
from pynput.mouse import Button, Controller as MouseController
from pynput.keyboard import Key, Listener, Controller as KeyboardController

mouse = MouseController()
keyboard = KeyboardController()

r = sr.Recognizer()
str = ''



# Class
class WindowsBalloonTip:
    def __init__(self, title, msg):
        message_map = { win32con.WM_DESTROY: self.OnDestroy,}

        # Register the window class.
        wc = WNDCLASS()
        hinst = wc.hInstance = GetModuleHandle(None)
        wc.lpszClassName = 'PythonTaskbar'
        wc.lpfnWndProc = message_map # could also specify a wndproc.
        classAtom = RegisterClass(wc)

        # Create the window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = CreateWindow(classAtom, "Taskbar", style, 0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, 0, 0, hinst, None)
        UpdateWindow(self.hwnd)

        # Icons managment
        iconPathName = os.path.abspath(os.path.join( sys.path[0], 'balloontip.ico' ))
        icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
        try:
            hicon = LoadImage(hinst, iconPathName, win32con.IMAGE_ICON, 0, 0, icon_flags)
        except:
            hicon = LoadIcon(0, win32con.IDI_APPLICATION)
        flags = NIF_ICON | NIF_MESSAGE | NIF_TIP
        nid = (self.hwnd, 0, flags, win32con.WM_USER+20, hicon, 'Tooltip')

        # Notify
        Shell_NotifyIcon(NIM_ADD, nid)
        Shell_NotifyIcon(NIM_MODIFY, (self.hwnd, 0, NIF_INFO, win32con.WM_USER+20, hicon, 'Balloon Tooltip', msg, 200, title))
        # self.show_balloon(title, msg)
        time.sleep(5)

        # Destroy
        DestroyWindow(self.hwnd)
        classAtom = UnregisterClass(classAtom, hinst)
    def OnDestroy(self, hwnd, msg, wparam, lparam):
        nid = (self.hwnd, 0)
        Shell_NotifyIcon(NIM_DELETE, nid)
        PostQuitMessage(0) # Terminate the app.

def balloon_tip(title, msg):
    w=WindowsBalloonTip(title, msg)
def balloon_tip2(title, msg):
    w2=WindowsBalloonTip(title, msg)




while('close' not in str and 'shut up' not in str):
    with sr.Microphone() as source:
        audio = r.adjust_for_ambient_noise(source)
        

        audio = r.listen(source, phrase_time_limit = 4)
        print('Say Anything')
        balloon_tip('Say Anything', 'Say it now!!!')

        try:
            text = r.recognize_google(audio)

            print("You said : {}".format(text))
            balloon_tip('Received', "You said : {}".format(text))

            str = format(text)


            if('scroll down' in str):
                mouse.scroll(0,-50)
                str = ''
            if('scroll up' in str):
                mouse.scroll(0,50)
                str = ''

                
            if('1, 1' in str):
                mouse.position = (160,80)
                str = ''
            if('1, 2' in str):
                mouse.position = (520,80)
                str = ''
            if('1, 3' in str):
                mouse.position = (880,80)
                str = ''
            if('1, 4' in str):
                mouse.position = (1240,80)
                str = ''
                
            if('2, 1' in str):
                mouse.position = (160,240)
                str = ''
            if('2, 2' in str):
                mouse.position = (520,240)
                str = ''
            if('2, 3' in str):
                mouse.position = (880,240)
                str = ''
            if('2, 4' in str):
                mouse.position = (1240,240)
                str = ''
                
            if('3, 1' in str):
                mouse.position = (160,400)
                str = ''
            if('3, 2' in str):
                mouse.position = (520,400)
                str = ''
            if('3, 3' in str):
                mouse.position = (880,400)
                str = ''
            if('3, 4' in str):
                mouse.position = (1240,400)
                str = ''
                
            if('4, 1' in str):
                mouse.position = (160,560)
                str = ''
            if('4, 2' in str):
                mouse.position = (520,560)
                str = ''
            if('4, 3' in str):
                mouse.position = (880,560)
                str = ''
            if('4, 4' in str):
                mouse.position = (1240,560)
                str = ''
                
            if('5, 1' in str):
                mouse.position = (160,720)
                str = ''
            if('5, 2' in str):
                mouse.position = (520,720)
                str = ''
            if('5, 3' in str):
                mouse.position = (880,720)
                str = ''
            if('5, 4' in str):
                mouse.position = (1240,720)
                str = ''

            if(('left click') in str or ('left-click') in str):
                mouse.click(Button.left,1)
                str = ''
            if(('right click') in str or ('right-click' in str)):
                mouse.click(Button.right,1)
                str = ''
            if(('double click' in str) or ('double-click' in str)):
                mouse.click(Button.left,2)
                str = ''

            '''Additional mouse command'''
            if ('open file' in str):
                mouse.click(Button.left, 3)
                str = ''

            if ('start button' in str):
                 mouse.position = (0,7320)
                 str = ''


            

            ''' Keyboard command start here!!!!!!!!!!!!!!!!!!!'''
            if (('zoom in') in str):
                pyautogui.keyDown('ctrl')
                pyautogui.keyDown('+')
                pyautogui.keyUp('+')
                pyautogui.keyUp('ctrl')
            if (('zoom out') in str):
                pyautogui.keyDown('ctrl')
                pyautogui.keyDown('-')
                pyautogui.keyUp('-')
                pyautogui.keyUp('ctrl')

            if ('type' in str):
                audio2 = r.adjust_for_ambient_noise(source)        
                print('Say what you would like to type')
                audio2 = r.listen(source, phrase_time_limit = 2)
                text2 = r.recognize_google(audio2)
                print("You said : {}".format(text2))
                str2 = format(text2)
                pyautogui.typewrite(str2)
            
            if ('enter' in str):
                pyautogui.keyDown('enter')
                pyautogui.keyUp('enter')
            
            if ('quit' in str):
                pyautogui.keyDown('ctrl')
                pyautogui.keyDown('w')
                pyautogui.keyUp('w')
                pyautogui.keyUp('ctrl')

            if ('volume down' in str):
                pyautogui.keyDown('f11')
                pyautogui.keyUp('f11')
                
                



        except:
            print("Sorry I could not recognize what you said")
            balloon_tip2('Message: ',"Sorry I could not recognize what you said")





    
    

    
    