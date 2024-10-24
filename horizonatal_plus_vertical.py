import win32com.client  # library for interacting with COM applications like PowerPoint
from cvzone.HandTrackingModule import HandDetector  # hand tracking module to detect hand gestures
import cv2  # OpenCV for capturing video from the webcam
import time  # to handle delays and timing
import numpy as np  # for numerical operations

# Initialize PowerPoint application
Application = win32com.client.Dispatch("PowerPoint.Application")
Presentation = Application.Presentations.Open("D:\PPT Controller\Sample.pptx")
Presentation.SlideShowSettings.Run()

# Parameters for webcam and gesture detection
width, height = 900, 720  # setting dimensions for the webcam feed
inputDelay = 1.5  # increased delay to ensure a gesture is recognized (in seconds)

cap = cv2.VideoCapture(0)  # capturing video from default/webcam
cap.set(3, width)  # setting the width of the frame
cap.set(4, height)  # setting the height of the frame

# Hand detector (detection confidence can be adjusted)
detectorHand = HandDetector(detectionCon=0.8, maxHands=1)

# Variables to manage gesture control
lastPosition = None  # storing the last position of the palm
lastGestureTime = time.time()  # tracking the last time a gesture was made

def is_slide_show_running():
    """Check if the slide show is running to prevent exceptions."""
    try:
        return Presentation.SlideShowWindow.View is not None
    except:
        return False

while True:
    # Capturing webcam frame
    success, img = cap.read()
    if not success:
        break

    # Detecting hand in the captured frame
    hands, img = detectorHand.findHands(img)

    if hands:
        hand = hands[0]
        palm = hand['center']  # Getting the center coordinates of the palm
        fingers = detectorHand.fingersUp(hand)  # Getting the state of the fingers (up or down)
        lmList = hand['lmList']  # Landmark list (21 points representing the hand)

        # Checking for punching gesture (all fingers down), then stopping the slideshow
        if fingers == [0, 0, 0, 0, 0]:
            print("Stop Slideshow")
            if is_slide_show_running():
                Presentation.SlideShowWindow.View.Exit()
            break

        # If this is the first detected hand position, initializing lastPosition
        if lastPosition is None:
            lastPosition = palm
            continue

        # Calculating the difference in x and y coordinates
        dx = palm[0] - lastPosition[0]
        dy = palm[1] - lastPosition[1]

        # Check if the hand is vertical or horizontal
        palm_tip_y = lmList[0][1]
        middle_finger_tip_y = lmList[12][1]

        is_hand_vertical = abs(palm_tip_y - middle_finger_tip_y) > 100

        # Get fingertips' x positions relative to the palm center
        fingertips = [lmList[i][0] for i in [8, 12, 16, 20]]
        palm_center_x = palm[0]

        # Relaxed condition: if at least 2 fingertips go left or right from the palm center
        fingertips_to_left = sum(fingertip < palm_center_x for fingertip in fingertips) >= 2
        fingertips_to_right = sum(fingertip > palm_center_x for fingertip in fingertips) >= 2

        if time.time() - lastGestureTime > inputDelay:
            if is_slide_show_running():  # Check if slide show is running
                if is_hand_vertical or not is_hand_vertical:  # Now responds to both vertical and horizontal hand positions
                    if fingertips_to_left:
                        print("Next Slide")
                        Presentation.SlideShowWindow.View.Next()
                        lastGestureTime = time.time()
                    elif fingertips_to_right:
                        print("Previous Slide")
                        Presentation.SlideShowWindow.View.Previous()
                        lastGestureTime = time.time()

        # Updating last position
        lastPosition = palm

    # Showing the webcam feed with any detected hands
    cv2.imshow("Look here", img)

    # Exit condition (press 'q' to quit)
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
