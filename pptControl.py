import win32com.client
from cvzone.HandTrackingModule import HandDetector
import cv2
import os
import numpy as np
import aspose.slides as slides
import aspose.pydrawing as drawing


def main():

     Application = win32com.client.Dispatch("PowerPoint.Application")
     Presentation = Application.Presentations.Open("C:\\Users\DELL\Downloads\MajorProject_Virtumotion.pptx")
     print(Presentation.Name)
     Presentation.SlideShowSettings.Run()
     # Parameters
     width, height = 900, 720
     gestureThreshold = 300
     # Camera Setup
     cap = cv2.VideoCapture(0)
     cap.set(3, width)
     cap.set(4, height)
     # Hand Detector
     detectorHand = HandDetector(detectionCon=0.8, maxHands=1)
     # Variables

     delay = 30
     buttonPressed = False
     counter = 0
     imgNumber = 20


     while True:
         # Get image frame
         success, img = cap.read()
         # Find the hand and its landmarks
         hands, img = detectorHand.findHands(img)  # with draw
         cv2.line(img, (0, gestureThreshold), (width, gestureThreshold), (0, 255, 0), 10)
         if hands and buttonPressed is False:  # If hand is detected
             hand = hands[0]
             cx, cy = hand["center"]
             lmList = hand["lmList"]  # List of 21 Landmark points

             # Constraints value for easier drawing
             xVal = int(np.interp(lmList[8][0], [width // 2, width], [0, width]))
             yVal = int(np.interp(lmList[8][1], [150, height - 150], [0, height]))
             indexFinger = xVal, yVal
             # cx, cy = lmList[8][:2]  # Tip of the index finger

             fingers = detectorHand.fingersUp(hand)  # List of which fingers are up
             if cy <= gestureThreshold:  # If hand is at the height of the face
                 if fingers == [1, 0, 0, 0, 0]:
                     print("Previous")
                     buttonPressed = True
                     if imgNumber > 0:
                         Presentation.SlideShowWindow.View.Previous()
                         imgNumber -= 1
                         annotations = [[]]
                         annotationNumber = -1
                         annotationStart = False
                 if fingers == [0, 0, 0, 0, 1]:
                     print("Next")
                     buttonPressed = True
                     if imgNumber > 0:
                         Presentation.SlideShowWindow.View.Next()
                         imgNumber += 1
                         annotations = [[]]
                         annotationNumber = -1
                         annotationStart = False

         if buttonPressed:
             counter += 1
             if counter > delay:
                 counter = 0
                 buttonPressed = False

         cv2.imshow("Image", img)
         key = cv2.waitKey(1)
         if key == ord('q'):
          break


if __name__ == "__main__":
    main()
