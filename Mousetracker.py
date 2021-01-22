import win32api
import time
import pyautogui as pa

state_left = win32api.GetKeyState(0x01)  # left button off=0 on =1
state_right = win32api.GetKeyState(0x02)  # right button off=0 on =1

while True:
    a = win32api.GetKeyState(0x01)
    b = win32api.GetKeyState(0x02)

    if a != state_left:  
        state_left = a
        if a < 0:
            print(f'Left Button- {pa.position()}')
        else:
            pass
    if b != state_right:  
        state_right = b
        
        if b < 0:
            print(f'Right Button- {pa.position()}')
        else:
            pass
        
    time.sleep(0.001)
