# Importing The Required Libraries
import cv2
import mediapipe as mp
import time
import handTrackingModule as htm
import math
import numpy as np
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# Creating Audio Driver Object
devices = AudioUtilities.GetSpeakers()

# Creating an Audio Interface
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

# Getting Volume Range
volRange = volume.GetVolumeRange() # In my case it was (-65.25, 0.0)

# Setting up independent variables for Minimum and Maximum Volume
minVolume = volRange[0]
maxVolume = volRange[1]
vol = 0

# pTime and cTime are for capturing Frames per Second (FPS)
pTime = 0
cTime = 0

# Initializing Video capturing object 
cap = cv2.VideoCapture(0)
cap.set(3, 1200)
cap.set(4, 720)

# Setting up Hand Detector Object
detector = htm.handDetector()

while True:
    # Capturing Frame
    _, img = cap.read()
    black = np.zeros(img.shape, np.uint8)
    img =  detector.findHands(img, img2=black)

    # Fetching LandMark List
    lmlist = detector.findPosition(img, draw=False)
    
    if lmlist:
        # Using Index and Thumb Fingers Locations in the image to get distance between them
        x1, y1 = lmlist[4][1], lmlist[4][2]
        x2, y2 = lmlist[8][1], lmlist[8][2]
        # Mid point of Thumb Tip and Index Tip
        cx, cy = (x1+x2)//2, (y1+y2)//2

        # Drawing Circles and Lines for Thumb and Index Finger Tips and a line between them
        cv2.circle(black, (x1,y1), 15, (209,206,0), cv2.FILLED)
        cv2.circle(black, (x2,y2), 15, (209,206,0), cv2.FILLED)
        cv2.line(black, (x1,y1), (x2,y2), (255,255,255), 3)

        # Finding Length between Thumb and Index Finger Tip
        length = math.hypot(x2-x1, y2-y1)

        # Making an interpretation of length in the range 0 to 125 for a total variable change of 240 units
        gradient = int((length/240)*125)
        # Drawing a circle at the mid point of the tip of the thumb and the index finger and gradually changing the color based on the volume level
        # from Grey (Minimum Volume i.e. Mute) to Red (Maximum Volume)
        cv2.circle(black, (cx,cy), 15, (125-gradient,125-gradient,125+gradient), cv2.FILLED)


        # Setting Volume level to interpreted result of distance between index and thumb finger tips

        # Here we are getting the interpretted volume level. Value of variable vol is set to minVolume if distance<=30 and maxVolume if distance>=270
        vol = int(np.interp(length, (30,270), (minVolume,maxVolume))) 
        # Making a trackBar volume variable for indicating visually the volume level. Interprettig it on a scale of (0,100) for a range (30,270)
        trackBarVolume = int(np.interp(length, (30,270), (0,100)))
        # Setting Master Volume Level
        volume.SetMasterVolumeLevel(vol, None)
        
    
    # Making a Visual Track bar indicator for volume level
    cv2.rectangle(black, (48, 239), (87,441), (209,206,0), 2)
    
    # Checking if trackBarVolume is defined. We need to perform this check because until a hand is detected,
    # the variable will not be declared and the rectangle function will give a 'variable not defined' error.
    # If the variable is defined, we can set the visual indicator to the appropriate level and fill the rectangle.
    if 'trackBarVolume' in vars(): 
        cv2.rectangle(black, (50,(220-trackBarVolume)*2), (85,220*2), (1,86,245), cv2.FILLED)


    # To calculate Frames Per Second
    cTime = time.time()
    fps = 1/(cTime - pTime)
    pTime = cTime

    # Putting FPS on the image
    cv2.putText(black, f'FPS -> {int(fps)}', (20,50), cv2.FONT_HERSHEY_COMPLEX_SMALL, 1, (124,71,200), 1, cv2.LINE_8)

    # Displaying the Frame
    cv2.imshow("Frame2", black)

    # Press the key 'q' to exit
    if cv2.waitKey(1) == ord('q'):
        cap.release()                                       # Releasing the Camera Object
        cv2.destroyAllWindows()                             # Closing Video Camera Window
        cv2.imshow('Thank You', cv2.imread('tk.png'))       # Viewing Thank You Message
        cv2.waitKey(1300)                                   # Wait for 2 Seconds
        break

# Destroying al Windows just to make sure and be Safe
cv2.destroyAllWindows()