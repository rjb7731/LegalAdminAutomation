import pyautogui 
import cv2 
import numpy as np

import win32api
import time
import pyautogui as pa

state_left = win32api.GetKeyState(0x01)  # left button off=0 on =1
state_right = win32api.GetKeyState(0x02) # right button off=0 on =1

# Specify resolution 
resolution = (1920, 1080) 

# Specify video codec 
codec = cv2.VideoWriter_fourcc(*"XVID") 

# Specify name of Output file 
filename = "Recording.avi"

# Specify frames rate. We can choose any 
# value and experiment with it 
fps = 15


# Creating a VideoWriter object 
out = cv2.VideoWriter(filename, codec, fps, resolution) 

# Create an Empty window 
cv2.namedWindow("Live", cv2.WINDOW_NORMAL) 

# Resize this window 
cv2.resizeWindow("Live", 480, 270)

#added in
pos = (911, 1045)
color = (0, 0, 255)

def getmouseposition():
    mousepos = pa.position()
    return mousepos[0],mousepos[1]



while True:
        a = win32api.GetKeyState(0x01)
        b = win32api.GetKeyState(0x02)
    # Take screenshot using PyAutoGUI 
        img = pyautogui.screenshot()
    # Convert the screenshot to a numpy array 
        frame = np.array(img)
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

    # Write it to the output file
        
    # Optional: Display the recording screen

        if a != state_left:
            state_left = a
            if a < 0:
                mouse_coordinates = getmouseposition()
                cv2.circle(frame, mouse_coordinates, 10, color, -1)
                cv2.putText(frame, f'Left Button- {pa.position()}', pos, cv2.FONT_HERSHEY_SIMPLEX,  
                   1, color, 2, cv2.LINE_AA)
                time.sleep(0.1)
            else:
                pass
        if b != state_right:
            state_left = b
            if a < 0:
                cv2.putText(frame, f'Right Button- {pa.position()}', pos, cv2.FONT_HERSHEY_SIMPLEX,  
                   1, color, 2, cv2.LINE_AA)
                time.sleep(1)
            else:
                pass
                
        out.write(frame)
        cv2.imshow('Live', frame)
        
                        
    # Stop recording when we press 'q'
        if cv2.waitKey(1) == ord('q'): 
            break

# Release the Video writer 
out.release() 

# Destroy all windows 
cv2.destroyAllWindows()
