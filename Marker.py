# Marker v1.0.0
# Nicholas Culmone
# A program that marks mutiple choice tests using OpenCV
# Created for Python Version 3.6.2 running on Windows 10

# Usage:

# When marking using the bubble sheet, hold the sheet sideways, with the
# RIGHT side facing upwards.

# When detecting the test, press the SPACE key when all of the coloured
# outlines are shown around the marker bubbles, the program will print out
# if the student was correct for every bubble and their total score of
# correct answers.

################################
# IMPORT STATEMENTS
################################
import cv2
import numpy as np
import argparse
import operator
import math
import ctypes
import sys
import os
import pandas as pd
import xlrd


################################
# GLOBAL VARIABLES FOLLOW
################################

numLen = 0
numWid = 0
answers = []
LETTER = ['A', 'B', 'C', 'D', 'E']

leftSide = []
rightSide = []
topSide = []
bottomSide = []


# Confgures the Blob Detector
params = cv2.SimpleBlobDetector_Params()
params.filterByArea = True
params.minArea = 150
params.filterByCircularity = True
params.minCircularity = 0.7


# Version detection for OpenCV and SimpleBlobDetector, creates Blob Detector
is_cv3 = cv2.__version__.startswith("3.")
if is_cv3:
    detector = cv2.SimpleBlobDetector_create(params)
else:
    detector = cv2.SimpleBlobDetector(params)


################################
# FUNCTION DEFINITIONS FOLLOW
################################


def goUntilMarked():
    global leftSide, rightSide, topSide, bottomSide

    cv2.namedWindow("Marker")
    vc = cv2.VideoCapture(0)

    path = os.path.dirname(os.path.realpath(__file__))

    
    # Gets the first frame, makes sure opened correctly
    if vc.isOpened():
        rval, frame = vc.read()
    else:
        rval = False

    try:
        os.mkdir(path + "\\Students")
    except OSError:
        pass

    path = path + "\\Students\\"


    while rval:
        # GrayScales the image
        tmp = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY);

        # Blur's the image so partially filled bubbles and background noise not detected
        frame_gray = blur = cv2.blur(tmp,(1,8))

        # Detects all of the filled bubbles
        keypoints = detector.detect(frame_gray)
        keypoints.sort(key=operator.attrgetter('pt'))

        # Gets the marker points
        leftSide = []
        rightSide = []
        topSide = []
        bottomSide = []

        # Finds which points are the outlines, and sorts them into arrays
        if len(keypoints) >= 2*numLen + 2*numWid:
            keypoints.sort(key=lambda x: x.pt[1])
            
            for i in range(0,numLen):
                leftSide.append(keypoints[0])
                keypoints.pop(0)
                
            for i in range(0,numLen):
                rightSide.append(keypoints[len(keypoints) - 1])
                keypoints.pop(len(keypoints) - 1)
                
            leftSide.sort(key=lambda x: x.pt[0])
            rightSide.sort(key=lambda x: x.pt[0])

            keypoints.sort(key=lambda x: x.pt[0])

            for i in range(0,numWid):
                topSide.append(keypoints[0])
                keypoints.pop(0)

            for i in range(0,numWid):
                bottomSide.append(keypoints[len(keypoints) - 1])
                keypoints.pop(len(keypoints) - 1)
                
            topSide.sort(key=lambda x: x.pt[1])
            bottomSide.sort(key=lambda x: x.pt[1])


        # Provides an image for the user (change tmp to frame once done)
        ansImg = cv2.drawKeypoints(tmp, keypoints, np.array([]), (0,0,255), cv2.DRAW_MATCHES_FLAGS_DRAW_RICH_KEYPOINTS)
        topAnsImg = cv2.drawKeypoints(ansImg, topSide, np.array([]), (255,0,0), cv2.DRAW_MATCHES_FLAGS_DRAW_RICH_KEYPOINTS)
        botAnsImg = cv2.drawKeypoints(topAnsImg, bottomSide, np.array([]), (255,255,0), cv2.DRAW_MATCHES_FLAGS_DRAW_RICH_KEYPOINTS)
        rightAnsImg = cv2.drawKeypoints(botAnsImg, rightSide, np.array([]), (255,0,255), cv2.DRAW_MATCHES_FLAGS_DRAW_RICH_KEYPOINTS)
        im_with_keypoints = cv2.drawKeypoints(rightAnsImg, leftSide, np.array([]), (0,255,0), cv2.DRAW_MATCHES_FLAGS_DRAW_RICH_KEYPOINTS)
        cv2.imshow("Marker", im_with_keypoints)

        rval, frame = vc.read()

        # When spacebar is pressed, all of the points are checked for their position
        key = cv2.waitKey(20)
        if key == 32 and len(topSide) > 0:

            ans = [[False for i in range(numWid)] for j in range(numLen)]
            for i in range(0, len(keypoints)):
                pointX, pointY = findPos(keypoints[i])
                ans[pointY][pointX] = True

            vc.release()
            cv2.destroyWindow("Marker")
            return ans
                        
            
        # If the window is destroyed or ESC is pressed, ends the function
        if key == 27 or cv2.getWindowProperty("Marker", 0) < 0:
            vc.release()
            cv2.destroyWindow("Marker")
            return 0


# Input: KeyPoint p
# Output: the position at which the dot point is located
def findPos(p):
    distSh = 100000
    xPos = 0
    yPos = 0

    for i in range(0,numWid):
        for j in range(0, numLen):
            x,y = findIntersection(leftSide[j].pt[0], leftSide[j].pt[1], rightSide[j].pt[0], rightSide[j].pt[1], topSide[i].pt[0], topSide[i].pt[1], bottomSide[i].pt[0], bottomSide[i].pt[1])
            dist = distance(p.pt[0], p.pt[1], x, y)
            if dist < distSh:
                distSh = dist
                xPos = i
                yPos = j

    return numWid - xPos - 1, yPos


# Finds the intersection of two lines
# Input: Coordinates of four points; the first two for the first line,
# the next two for the next line
def findIntersection(x1,y1,x2,y2,x3,y3,x4,y4):
    b1 = y2 - (x2*((y2-y1)/(x2-x1)))
    b2 = y4 - (x4*((y4-y3)/(x4-x3)))
    m1 = (y2-y1)/(x2-x1)
    m2 = (y4-y3)/(x4-x3)

    x = (b2-b1)/(m1-m2)
    y = (b1*m2 - b2*m1)/(m2-m1)

    return x,y


# Distance formula
# Input: Coordinates for two points
def distance(x1,y1,x2,y2):
    dist = math.sqrt((x2 - x1)**2 + (y2 - y1)**2)
    return dist


# Marks a Sheet with every bubble being treated as an individual question
# Input: Answer array, master answer sheet
def everyBubbleMarked(ans, master):
    correct = [[False for i in range(numWid)] for j in range(numLen)]
    score = 0
    
    for i in range(0, numWid):
        for j in range(0, numLen):
            if ans[j][i] == master[j][i]:
                correct[j][i] = True
                score += 1
                
    return correct, score


# Goes through an Excel file to get each student, and then get a mark for them
# Input: name of Excel file
def readFiles(fileName):
    df = pd.read_excel(fileName + ".xlsx", header = None)
    df.columns = ["NAME", "SID"]

    mark=[]
    for index, row in df.iterrows():
        print("\nName: ", row['NAME'], "\nSID: ", row['SID'])
        
        ans = goUntilMarked()
        if ans == 0:
            mark.append(0)
        else:
            correct, totalGrade = everyBubbleMarked(ans, answers)
            print("Score: ", totalGrade)
            mark.append(totalGrade)
            exportToFile(row['SID'], ans, totalGrade)
        
    df["MARK"]= mark
    df.to_excel(fileName + "Out.xlsx")


# Exports the results of a student's test into a csv file
# Input: Student Number, Student's answers, Total grade
def exportToFile(stuName, ans, totalGrade):
    path = os.path.dirname(os.path.realpath(__file__))
    path = path + "\\Students\\"
    pathName = path + "" + str(stuName) + ".csv"

    with open(pathName,"w+") as f:
        f.write(",")
        for i in range(0, numWid):
            f.write(LETTER[i] + ",")
        f.write("\n")
        
        for i in range (0, len(ans)):
            f.write(str(i+1) + ",")
            for j in range(0, len(ans[i])):
                if ans[i][j] == True:
                    f.write("T")
                    
                if j != len(ans[i]) - 1:
                    f.write(",")

            if i != len(ans) - 1:
                f.write("\n")
        f.write("\n\nGrade," + str(totalGrade))



################################
# MAIN PROGRAM FOLLOWS
################################

stuNums = []
grades = []

while True:
    print("\n\nSelect an Option:\n",
          "0. Exit\n",
          "1. Enter an Answer Key\n",
          "2. Mark Individual Tests\n",
          "3. Mark Tests Based on students.txt\n\n")

    option = input("Enter Selection: ")

    if option == '0':
        path = os.path.dirname(os.path.realpath(__file__))

        with open(path + "\\studentsOut.csv" , "w+") as f:
            for i in range(0, len(stuNums)):
                f.write(str(stuNums[i]) + "," + str(grades[i]) + "\n")

        sys.exit(0)
    elif option == '1':
        numLen = int(input("\nEnter number of questions: "))
        numWid = int(input("Enter number of options per question: "))

        answers = goUntilMarked()

        print("\n\nCorrect Answers:")
        for i in range (0, numLen):
            print("\nQuestion",i+1,": ", end="")
            for j in range(0, numWid):
                print(answers[i][j], "\t", end="")

        input("\n\nPress Enter to Continue")
        

    elif option == '2':
        if answers == []:
            input("\n\nThere is no answer key currently\nPress Enter to Continue")
        else:
            print("\nPress ESC or close graphics window to stop marking...")
            while True:
                ans = goUntilMarked()
                if ans == 0:
                    break

                correct, totalGrade = everyBubbleMarked(ans, answers)
                stuNum = input("Enter student number: ")

                stuNums.append(stuNum)
                grades.append(totalGrade)

                exportToFile(stuNum, ans, totalGrade)

    elif option == '3':
        readFiles("students")

    else:
        input("\n\nInvalid selection\nPress Enter to Continue")