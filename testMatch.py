import pywinauto
import win32gui
import cv2
import numpy as np
import time
import random
import sys

from ctypes import *
from matplotlib import pyplot as plt
from mss import mss
from PIL import Image

APP_PATH = "C:\\Program Files\\Nox\\bin\\Nox.exe"
UI_WIDTH_720 = 1280
SHOW_MATCH = True
APPBAR_H = 32
CONSOLE_SLEEP_TIME = 0.2
LOADING_DOT = '.'
BACKSPACE = '\b \b'
TESTING_MODE = True
fuckOnagda = True
fuckBen = True
ADJENCENCY_THRESHOLD = 0.5

class Vision:
    def __init__(self):
        self.enemy_templates = {
            'BigBS'       : 'assets/BigBS.png',
            'MedBS'       : 'assets/MedBS.png',
            'BigCrus'     : 'assets/BigCrus.png',
            'MedCrus'     : 'assets/MedCrus.png',
            'BigMoney'    : 'assets/BigMoney.png'
        }
        self.non_enemy_templates = {
            'T4-Box'      : 'assets/T4-Box.png',
            'Gear'        : 'assets/Gear.png',
            'startBattle' : 'assets/BattleStart.png',
            'endBattle'   : 'assets/BtFin.png',
            'confirm'     : 'assets/Confirm.png',
            'switch'      : 'assets/switch.png'
        }
        self.empty_tile_templates = {
            'PLACEHOLDER' : 'assets/PLACEHOLDER.png'
        }
        self.landmarks = {
            'Iceberg'     : 'assets/Iceberg.png',
            'Command'     : 'assets/Command.png'
        }

        self.enemy_templates = { k: cv2.imread(v, 0) for (k, v) in self.enemy_templates.items() }
        self.nonEnemyTemplates = { k: cv2.imread(v, 0) for (k, v) in self.non_enemy_templates.items() }
        self.emptyTileTemplates = { k: cv2.imread(v, 0) for (k, v) in self.empty_tile_templates.items() }
        self.landmarks = { k: cv2.imread(v, 0) for (k, v) in self.landmarks.items() }
        self.screen = mss()
        self.frame = None

    def setMonitor(self, monitorDimensions):
        self.monitor = {'top': monitorDimensions['top'], 'left': monitorDimensions['left'], 'width': monitorDimensions['width'], 'height': monitorDimensions['height']}

    def convert_rgb_to_bgr(self, img):
        return img[:, :, ::-1]

    def take_screenshot(self):
        #print (self.monitor)
        sct_img = self.screen.grab(self.monitor)
        img = Image.frombytes('RGB', sct_img.size, sct_img.rgb)
        img = np.array(img)
        img = self.convert_rgb_to_bgr(img)
        img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        return img_gray

    def refresh_frame(self):
        self.frame = self.take_screenshot()

    def save_screenshot(self):
        img = self.take_screenshot()
        cv2.imwrite('assets/test1.png', img)

    def match_template(self, img_grayscale, template, threshold):
        """
        Matches template image in a target grayscaled image
        """

        res = cv2.matchTemplate(img_grayscale, template, cv2.TM_CCOEFF_NORMED)
        matches = np.where(res >= threshold)
        return matches

    def scaled_find_template(self, name, threshold, templateSet, scales, image=None):
        if image is None:
            if self.frame is None:
                self.refresh_frame()

            image = self.frame

        # try catch for KeyError and each error try a different array
        if(templateSet == 'ENEMY'):
            initial_template = self.enemy_templates[name]
        elif(templateSet == 'EMPTY'):
            initial_template = self.emptyTileTemplates[name]
        elif(templateSet == 'LANDMARKS'):
            initial_template = self.landmarks[name]
        else:
            initial_template = self.nonEnemyTemplates[name]

        for scale in scales:
            scaled_template = cv2.resize(initial_template, (0,0), fx=scale, fy=scale)
            matches = self.match_template(
                image,
                scaled_template,
                threshold
            )
            if np.shape(matches)[1] >= 1:
                if(SHOW_MATCH == True):
                    w, h = scaled_template.shape[::-1]
                    for pt in zip(*matches[::-1]):
                        cv2.rectangle(image, pt, (pt[0] + w, pt[1] + h), (0,255,255), 1)
                        midpointX = int(pt[0] + (0.5*w))
                        midpointY = int(pt[1] + (0.5*h) + APPBAR_H)
                    cv2.imshow('image',image)
                    cv2.waitKey(0)
                    cv2.destroyAllWindows()
                return matches
        return matches

    def find_template(self, name, threshold, image=None):
        if image is None:
            if self.frame is None:
                self.refresh_frame()

            image = self.frame
        # Something bad about self.templates
        return self.match_template(
            image,
            self.templates[name],
            threshold
        )

    def find_scale(self, currentWidth):
        scale = currentWidth / UI_WIDTH_720
        return scale

def getWindowObject(appPath):
    noxAppDict = {}
    app = pywinauto.Application().connect(path=appPath)
    #window = app.top_window()
    allElements = app.window(title_re="NoxPlayer.*") #returns window spec object
    childWindow = allElements.child_window(title="ScreenBoardClassWindow")
    noxAppDict['allElements'] = allElements
    noxAppDict['mainWindow'] = childWindow
    return noxAppDict

def bringNoxToFront(noxApp):
    #bring window into foreground
    userDict = storeUserState()
    noxApp.set_focus()
    restoreUserState(userDict)

def getWindowDimensions(noxWindow):
    w = noxWindow.rectangle().width()
    h = noxWindow.rectangle().height()
    x = noxWindow.rectangle().left
    y = noxWindow.rectangle().top

    # print ("top left co-ords:" + str(x) + 'x' + str(y))
    # print ("size:" + str(w) + 'x' + str(h))

    windowObj = {'top': y, 'left': x, 'width': w, 'height': h}
    return windowObj

def getHWND():
    hwnd = pywinauto.win32functions.GetForegroundWindow()
    return hwnd

def storeUserState():
    userDict = {}
    userMouseX, userMouseY = win32gui.GetCursorPos()
    currWindowHWND = pywinauto.win32functions.GetForegroundWindow()
    userDict['userMouseX'] = userMouseX
    userDict['userMouseY'] = userMouseY
    userDict['userWindowHWND'] = currWindowHWND
#    print ('Saving usermouse at X: ' + str(userDict['userMouseX']) + ' Y: ' + str(userDict['userMouseY']))
    return userDict

def restoreUserState(userDict):
    currWindowHWND = pywinauto.win32functions.GetForegroundWindow()
#    print ('Restoring usermouse to X: ' + str(userDict['userMouseX']) + ' Y: ' + str(userDict['userMouseY']))
    pywinauto.win32functions.SetForegroundWindow(userDict['userWindowHWND'])
    pywinauto.mouse.move(coords=(userDict['userMouseX'], userDict['userMouseY']))

def checkPoints(newPoint, matched, w, h):
    isNewPoint = True
    tempX = []

    for point in matched:
        if (abs(newPoint[0] - point[0]) < w):
            if(abs(newPoint[1] - point[1]) < h):
                isNewPoint = False
    return isNewPoint

def checkPriority(newPoint, priorityAreas, w, h):
    isImportantShip = False
    if(abs(newPoint[0] - priorityAreas[0]) < w * ADJENCENCY_THRESHOLD):
        isImportantShip = True
    if(abs(newPoint[1] - priorityAreas[1]) < h * ADJENCENCY_THRESHOLD):
        isImportantShip = True

    if(isImportantShip == True):
        return newPoint
    else:
        return None

def matchTemplate(noxWindowDimensions, templateName, matchPercentage, templateSet):
    returnDict = {}

    vision = Vision()
    vision.setMonitor(noxWindowDimensions)

    UIscale = vision.find_scale(noxWindowDimensions['width'])
    scales = [UIscale]

    # try catch for KeyError and each error try a different array
    if(templateSet == 'ENEMY'):
        initial_template = vision.templates[templateName]
    elif(templateSet == 'EMPTY'):
        initial_template = vision.emptyTileTemplates[templateName]
    else:
        initial_template = vision.nonEnemyTemplates[templateName]

    scaled_template = cv2.resize(initial_template, (0,0), fx=UIscale, fy=UIscale)
    w, h = scaled_template.shape[::-1]
    matches = vision.scaled_find_template(templateName, matchPercentage, templateSet, scales=scales)

    returnDict['matches'] = matches
    returnDict['width'] = w
    returnDict['height'] = h
    return returnDict

def findPrioritySpace():
    tempMatches = []
    priorityAreas = []
    vision = Vision()
    vision.setMonitor(noxWindowDimensions)

    UIscale = vision.find_scale(noxWindowDimensions['width'])
    scales = [UIscale]

    for template in vision.landmarks:
        initial_template = vision.landmarks[template]
        scaled_template = cv2.resize(initial_template, (0,0), fx=UIscale, fy=UIscale)
        w, h = scaled_template.shape[::-1]

        matches = vision.scaled_find_template(template, 0.55, 'LANDMARKS', scales=scales)

        for pt in zip(*matches[::-1]):
            # add the template size to point
            realPointX = pt[0] + int(w * 0.5) + noxWindowDimensions['left']
            realPointY = pt[1] + int(h * 0.5) + noxWindowDimensions['top']
            newRealPoint = (realPointX,realPointY)

            if(len(tempMatches) >= 1):
                isNewPoint = checkPoints(newRealPoint, tempMatches, w, h)
                if (isNewPoint == True):
                    tempMatches.append(newRealPoint)
            else: #first point
                tempMatches.append(newRealPoint)

    if(tempMatches[0][0] < tempMatches[1][0]):
        priorityAreas.append(tempMatches[0][0] - 2 * (int(w * 0.5)))
        priorityAreas.append(tempMatches[1][1])
    else:
        priorityAreas.append(tempMatches[1][0] - 2 * (int(w * 0.5)))
        priorityAreas.append(tempMatches[0][1])

    return priorityAreas

def matchShips(priorityAreas):
    tempMatches = []
    priorityShips = []
    vision = Vision()
    vision.setMonitor(noxWindowDimensions)

    UIscale = vision.find_scale(noxWindowDimensions['width'])
    scales = [UIscale]

    for template in vision.enemy_templates:
        initial_template = vision.enemy_templates[template]
        scaled_template = cv2.resize(initial_template, (0,0), fx=UIscale, fy=UIscale)
        w, h = scaled_template.shape[::-1]

        matches = vision.scaled_find_template(template, 0.45, 'ENEMY', scales=scales)

        for pt in zip(*matches[::-1]):
            # add the template size to point
            realPointX = pt[0] + int(w * 0.5) + noxWindowDimensions['left']
            realPointY = pt[1] + int(h * 0.5) + noxWindowDimensions['top']
            newRealPoint = (realPointX,realPointY)

            if(len(tempMatches) >= 1):
                isNewShip = checkPoints(newRealPoint, tempMatches, w, h)
                isNewPriorityShip = checkPoints(newRealPoint, priorityShips, w, h)
                if (isNewShip == True and isNewPriorityShip == True):
                    isImportantShip = checkPriority(newRealPoint, priorityAreas, w, h)
                    if(isImportantShip == None):
                        tempMatches.append(newRealPoint)
                    else:
                        priorityShips.append(newRealPoint)
            else: #first point
                isNewShip = checkPoints(newRealPoint, tempMatches, w, h)
                isNewPriorityShip = checkPoints(newRealPoint, priorityShips, w, h)
                if (isNewShip == True and isNewPriorityShip == True):
                    isImportantShip = checkPriority(newRealPoint, priorityAreas, w, h)
                    if(isImportantShip == None):
                        tempMatches.append(newRealPoint)
                    else:
                        priorityShips.append(newRealPoint)

    return tempMatches, priorityShips

def chooseEnemy(matched):
    randNumb = random.randint(0, len(matched)-1)
    print ('Picked: ' + str(matched[randNumb]))
    userDict = storeUserState()
    ok = windll.user32.BlockInput(True) #enable block
    pywinauto.mouse.click(coords=(matched[randNumb][0],matched[randNumb][1]))
    time.sleep(0.5)
    ok = windll.user32.BlockInput(False) #disable block
    restoreUserState(userDict)
    del matched[randNumb]
    startBattle()
    inBattle()
    endBattle()
    return matched

def switchFleet():
    templateArray = matchTemplate(noxWindowDimensions, 'switch', 0.70, 'UI')
    points = []

    for pt in zip(*templateArray['matches'][::-1]):
        # add the template size to point
        realPointX = pt[0] + int(templateArray['width']*0.5) + noxWindowDimensions['left']
        realPointY = pt[1] + int(templateArray['height']*0.5) + noxWindowDimensions['top']
        newRealPoint = (realPointX,realPointY)

        if(len(points) >= 1):
            isNewPoint = checkPoints(newRealPoint, points, templateArray['width'], templateArray['height'])
            if (isNewPoint == True):
                points.append(newRealPoint)
        else:
            points.append(newRealPoint)
    if(len(points) > 0):
        userDict = storeUserState()
        ok = windll.user32.BlockInput(True) #enable block
        pywinauto.mouse.click(coords=(points[0]))
        ok = windll.user32.BlockInput(False) #disable block
        restoreUserState(userDict)

def chooseBoss(noxWindowDimensions):
    points = []
    # this finds an empty tile to click
    # Find midpoint of screen
    # Left click mouse, drag down to bottom of screen, let go of mouse
    templateArray = matchTemplate(noxWindowDimensions, 'city', 0.70, 'EMPTY')

    for pt in zip(*templateArray['matches'][::-1]):
        # add the template size to point
        realPointX = pt[0] + int(templateArray['width']*0.5) + noxWindowDimensions['left']
        realPointY = pt[1] + int(templateArray['height']*0.5) + noxWindowDimensions['top']
        newRealPoint = (realPointX,realPointY)

        if(len(points) >= 1):
            isNewPoint = checkPoints(newRealPoint, points, templateArray['width'], templateArray['height'])
            if (isNewPoint == True):
                points.append(newRealPoint)
        else:
            points.append(newRealPoint)
    if(len(points) > 0):
        # the point you click
        # when you drag down
        # use the x from the point you found the city at
        userDict = storeUserState()
        ok = windll.user32.BlockInput(True) #enable block
        pywinauto.mouse.press(button='left', coords=(points[0]))
        ok = windll.user32.BlockInput(False) #disable block
        time.sleep(0.1)
        ok = windll.user32.BlockInput(True) #enable block
        pywinauto.mouse.move(coords=(points[0][0],(noxWindowDimensions['top'] + int(noxWindowDimensions['height']))))
        ok = windll.user32.BlockInput(False) #disable block
        time.sleep(0.1)
        ok = windll.user32.BlockInput(True) #enable block
        pywinauto.mouse.release(button='left', coords=(points[0][0],noxWindowDimensions['top'] + int(noxWindowDimensions['height'])))
        ok = windll.user32.BlockInput(False) #disable block
        restoreUserState(userDict)
        points.clear()

    # This finds the boss
    # templateArray = matchTemplate(noxWindowDimensions, 'giantKizuna', 0.70, 'UI')

    for pt in zip(*templateArray['matches'][::-1]):
        # add the template size to point
        realPointX = pt[0] + int(templateArray['width']*0.5) + noxWindowDimensions['left']
        realPointY = pt[1] + int(templateArray['height']*0.5) + noxWindowDimensions['top']
        newRealPoint = (realPointX,realPointY)

        if(len(points) >= 1):
            isNewPoint = checkPoints(newRealPoint, points, templateArray['width'], templateArray['height'])
            if (isNewPoint == True):
                points.append(newRealPoint)
        else:
            points.append(newRealPoint)
    if(len(points) > 0):
        userDict = storeUserState()
        ok = windll.user32.BlockInput(True) #enable block
        pywinauto.mouse.click(coords=(points[0]))
        ok = windll.user32.BlockInput(False) #disable block
        restoreUserState(userDict)
        startBattle()
        inBattle()
        endBattle()

def startBattle():
    templateArray = matchTemplate(noxWindowDimensions, 'startBattle', 0.70, 'UI')

    points = []
    moving = True
    sys.stdout.write('Moving')
    while(moving == True):
        consolePrint()
        for pt in zip(*templateArray['matches'][::-1]):
            # add the template size to point
            realPointX = pt[0] + int(templateArray['width']*0.5) + noxWindowDimensions['left']
            realPointY = pt[1] + int(templateArray['height']*0.5) + noxWindowDimensions['top']
            newRealPoint = (realPointX,realPointY)

            if(len(points) >= 1):
                isNewPoint = checkPoints(newRealPoint, points, templateArray['width'], templateArray['height'])
                if (isNewPoint == True):
                    points.append(newRealPoint)
            else:
                points.append(newRealPoint)

        if(len(points) > 0):
            moving = False

        templateArray = matchTemplate(noxWindowDimensions, 'startBattle', 0.70, 'UI')

    print ('\nEntered fleet manager')
    #TODO make sure auto is checked here
    userDict = storeUserState()
    ok = windll.user32.BlockInput(True) #enable block
    pywinauto.mouse.click(coords=(points[0]))
    ok = windll.user32.BlockInput(False) #disable block
    restoreUserState(userDict)

def inBattle():
    templateArray = matchTemplate(noxWindowDimensions, 'endBattle', 0.70, 'UI')

    points = []
    battling = True
    sys.stdout.write('Battling')

    while(battling == True):
        consolePrint()
        userDict = storeUserState()
        for pt in zip(*templateArray['matches'][::-1]):
            realPointX = pt[0] + int(templateArray['width']*0.5) + noxWindowDimensions['left']
            realPointY = pt[1] + int(templateArray['height']*0.5) + noxWindowDimensions['top']
            newRealPoint = (realPointX,realPointY)

            if(len(points) >= 1):
                isNewPoint = checkPoints(newRealPoint, points, templateArray['width'], templateArray['height'])
                if (isNewPoint == True):
                    points.append(newRealPoint)
            else:
                points.append(newRealPoint)

        if(len(points) > 0):
            battling = False

        templateArray = matchTemplate(noxWindowDimensions, 'endBattle', 0.70, 'UI')
        time.sleep(0.5)

    print ('\nBattle ended')
    ok = windll.user32.BlockInput(True) #enable block
    pywinauto.mouse.click(coords=(points[0]))
    ok = windll.user32.BlockInput(False) #disable block
    restoreUserState(userDict)
    #TODO check next point incase of elite ship; reason: longer claim duration and extra click needed
    time.sleep(4)
    userDict = storeUserState()
    ok = windll.user32.BlockInput(True) #enable block
    pywinauto.mouse.click(coords=(points[0]))
    ok = windll.user32.BlockInput(False) #disable block
    restoreUserState(userDict)
    time.sleep(3)
    userDict = storeUserState()
    ok = windll.user32.BlockInput(True) #enable block
    pywinauto.mouse.click(coords=(points[0]))
    ok = windll.user32.BlockInput(False) #disable block
    restoreUserState(userDict)

def endBattle():
    templateArray = matchTemplate(noxWindowDimensions, 'confirm', 0.70, 'UI')

    points = []
    for pt in zip(*templateArray['matches'][::-1]):
        # add the template size to point
        realPointX = pt[0] + int(templateArray['width']*0.5) + noxWindowDimensions['left']
        realPointY = pt[1] + int(templateArray['height']*0.5) + noxWindowDimensions['top']
        newRealPoint = (realPointX,realPointY)

        if(len(points) >= 1):
            isNewPoint = checkPoints(newRealPoint, points, templateArray['width'], templateArray['height'])
            if (isNewPoint == True):
                points.append(newRealPoint)

        else:
            points.append(newRealPoint)
            userDict = storeUserState()
    ok = windll.user32.BlockInput(True) #enable block
    pywinauto.mouse.click(coords=(points[0]))
    ok = windll.user32.BlockInput(False) #disable block
    restoreUserState(userDict)
    time.sleep(5)

def consolePrint():
    for x in range(0, 3):
        sys.stdout.write(LOADING_DOT)
        sys.stdout.flush()
        time.sleep(CONSOLE_SLEEP_TIME)

    for x in range(0, 3):
        sys.stdout.write(BACKSPACE)
        sys.stdout.flush()
        time.sleep(CONSOLE_SLEEP_TIME)

userDict = storeUserState()

noxDict = getWindowObject(APP_PATH)
bringNoxToFront(noxDict['allElements'])
noxWindowDimensions = getWindowDimensions(noxDict['mainWindow'])
noxHWND = getHWND()

if(TESTING_MODE == False):
    matched = matchShips()

    if(len(matched) >= 4):
        for x in range(0, 4):
            matched = chooseEnemy(matched)
        switchFleet()
        chooseBoss(noxWindowDimensions)
    else:
        print ('Error less than 5 boats selected')

    print('End of script, following array should be left over enemies:')
    print(matched)
else:
    #print('TESTING MODE - YOU SHOULDN\'T BE HERE')
    priorityAreas = findPrioritySpace()
    ships, priorityShips = matchShips(priorityAreas)
    print('Found enemies at these positions: ')
    print(ships)
    print('Found priorityShips at: ')
    print(priorityShips)
