import pywinauto
import win32gui
import cv2
import numpy as np
import time
import random
import sys

from matplotlib import pyplot as plt
from mss import mss
from PIL import Image

APP_PATH = "D:\\Program Files\\Nox\\bin\\Nox.exe"
UI_WIDTH_720 = 1280
SHOW_MATCH = False
CONSOLE_SLEEP_TIME = 0.2
LOADING_DOT = '.'
BACKSPACE = '\b \b'
isKizunaSP4 = True

class Vision:
    def __init__(self):
        self.static_templates = {
            'Kizuna1'     : 'assets/KZ1.png',
            'Kizuna2'     : 'assets/KZ2.png',
            'Kizuna3'     : 'assets/KZ3.png'
        }
        self.non_enemy_static_templates = {
            'T4-Box'      : 'assets/T4-Box.png',
            'Gear'        : 'assets/Gear.png',
            'startBattle' : 'assets/BattleStart.png',
            'endBattle'   : 'assets/BtFin.png',
            'confirm'     : 'assets/Confirm.png',
            'giantKizuna' : 'assets/giantKizuna.png',
            'switch'      : 'assets/switch.png'
        }
        self.empty_tile_templates = {
            'empty'       : 'assets/Empty.png'
        }

        self.templates = { k: cv2.imread(v, 0) for (k, v) in self.static_templates.items() }
        self.nonEnemyTemplates = { k: cv2.imread(v, 0) for (k, v) in self.non_enemy_static_templates.items() }
        self.emptyTileTemplates = { k: cv2.imread(v, 0) for (k, v) in self.empty_tile_templates.items() }
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
            initial_template = self.templates[name]
        elif(templateSet == 'EMPTY'):
            initial_template = self.emptyTileTemplates[name]
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
                        midpointY = int(pt[1] + (0.5*h)+30)
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
        return self.match_template(
            image,
            self.templates[name],
            threshold
        )

    def find_scale(self, currentWidth):
        scale = currentWidth / UI_WIDTH_720
        return scale

def getWindowObject(appPath):
    app = pywinauto.Application().connect(path=appPath)
    #window = app.top_window()
    allElements = app.window(title_re="NoxPlayer.*") #returns window spec object
    childWindow = allElements.child_window(title="ScreenBoardClassWindow")
    return childWindow

def bringAppToFront(appWindow):
    #bring window into foreground
    if appWindow.has_style(pywinauto.win32defines.WS_MINIMIZE): # if minimized
        pywinauto.win32functions.ShowWindow(appWindow.wrapper_object(), 9) # restore window state
    else:
        pywinauto.win32functions.SetForegroundWindow(appWindow.wrapper_object()) #bring to front

def getWindowDimensions(appWindow):
    w = appWindow.rectangle().width()
    h = appWindow.rectangle().height()
    x = appWindow.rectangle().left
    y = appWindow.rectangle().top

    #TODO force window into certain dimensions before -30 pixel operation

    # print ("top left co-ords:" + str(x) + 'x' + str(y))
    # print ("size:" + str(w) + 'x' + str(h))

    windowObj = {'top': y, 'left': x, 'width': w, 'height': h}
    return windowObj

def returnFocus():
    print ('stub')

def focusLoop():
    print('stub')

def checkPoints(newPoint, matched, w, h):
    isNewPoint = True
    tempX = []

    #check why 1000, 400 isn't in array
    for point in matched:
        if (abs(newPoint[0] - point[0]) < w):
            if(abs(newPoint[1] - point[1]) < h):
                isNewPoint = False
    return isNewPoint

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

def initialMatch():
    tempMatches = []
    vision = Vision()
    vision.setMonitor(noxWindowDimensions)

    UIscale = vision.find_scale(noxWindowDimensions['width'])
    scales = [UIscale]

    for template in vision.templates:
        initial_template = vision.templates[template]
        scaled_template = cv2.resize(initial_template, (0,0), fx=UIscale, fy=UIscale)
        w, h = scaled_template.shape[::-1]

        matches = vision.scaled_find_template(template, 0.50, 'ENEMY', scales=scales)

        for pt in zip(*matches[::-1]):
            # add the template size to point
            realPointX = pt[0] + int(w * 0.5)
            realPointY = pt[1] + int(h * 1.5)
            newRealPoint = (realPointX,realPointY)

            if(len(tempMatches) >= 1):
                isNewPoint = checkPoints(newRealPoint, tempMatches, w, h)
                if (isNewPoint == True):
                    tempMatches.append(newRealPoint)
            else: #first point
                tempMatches.append(newRealPoint)

    print('Found enemies at these positions: ')
    print(tempMatches)
    return tempMatches

def chooseEnemy(matched):
    randNumb = random.randint(0, len(matched)-1)
    print ('Picked: ' + str(matched[randNumb]))
    pywinauto.mouse.click(coords=(matched[randNumb][0],matched[randNumb][1]))
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
        realPointX = pt[0] + int(templateArray['width']*0.5)
        realPointY = pt[1] + int(templateArray['height'])
        newRealPoint = (realPointX,realPointY)

        if(len(points) >= 1):
            isNewPoint = checkPoints(newRealPoint, points, templateArray['width'], templateArray['height'])
            if (isNewPoint == True):
                points.append(newRealPoint)
        else:
            points.append(newRealPoint)
    if(len(points) > 0):
        pywinauto.mouse.click(coords=(points[0]))

def chooseBoss(noxWindowDimensions):
    points = []

    # this finds an empty tile to click
    # Find midpoint of screen
    # Left click mouse, drag down to bottom of screen, let go of mouse
    templateArray = matchTemplate(noxWindowDimensions, 'empty', 0.70, 'EMPTY')
    for pt in zip(*templateArray['matches'][::-1]):
        # add the template size to point
        realPointX = pt[0] + int(templateArray['width']*0.5)
        realPointY = pt[1] + int(templateArray['height']*1.5)
        newRealPoint = (realPointX,realPointY)

        if(len(points) >= 1):
            isNewPoint = checkPoints(newRealPoint, points, templateArray['width'], templateArray['height'])
            if (isNewPoint == True):
                points.append(newRealPoint)
        else:
            points.append(newRealPoint)
    if(len(points) > 0):
        pywinauto.mouse.click(coords=(points[0]))
        pywinauto.mouse.press(button='left', coords=(int(noxWindowDimensions['width']/2),int(noxWindowDimensions['height']/2)))
        time.sleep(2)
        pywinauto.mouse.move(coords=(int(noxWindowDimensions['width']/2),int(noxWindowDimensions['height'])))
        time.sleep(2)
        pywinauto.mouse.release(button='left', coords=(int(noxWindowDimensions['width']/2),int(noxWindowDimensions['height'])))
        points.clear()

    # This finds the boss
    templateArray = matchTemplate(noxWindowDimensions, 'giantKizuna', 0.70, 'UI')

    for pt in zip(*templateArray['matches'][::-1]):
        # add the template size to point
        realPointX = pt[0] + int(templateArray['width']*0.5)
        realPointY = pt[1] + int(templateArray['height']*1.5)
        newRealPoint = (realPointX,realPointY)

        if(len(points) >= 1):
            isNewPoint = checkPoints(newRealPoint, points, templateArray['width'], templateArray['height'])
            if (isNewPoint == True):
                points.append(newRealPoint)
        else:
            points.append(newRealPoint)
    if(len(points) > 0):
        pywinauto.mouse.click(coords=(points[0]))
        time.sleep(9)
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
            realPointX = pt[0] + int(templateArray['width']*0.5)
            realPointY = pt[1] + int(templateArray['height']*0.5)
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
    pywinauto.mouse.click(coords=(points[0]))

def inBattle():
    templateArray = matchTemplate(noxWindowDimensions, 'endBattle', 0.70, 'UI')

    points = []
    battling = True
    sys.stdout.write('Battling')
    while(battling == True):
        consolePrint()
        for pt in zip(*templateArray['matches'][::-1]):
            realPointX = pt[0] + int(templateArray['width']*0.5)
            realPointY = pt[1] + int(templateArray['height']*1.5)
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
    pywinauto.mouse.click(coords=(points[0]))
    time.sleep(5)
    pywinauto.mouse.click(coords=(points[0]))
    time.sleep(3)
    pywinauto.mouse.click(coords=(points[0]))

def endBattle():
    templateArray = matchTemplate(noxWindowDimensions, 'confirm', 0.70, 'UI')

    points = []
    for pt in zip(*templateArray['matches'][::-1]):
        # add the template size to point
        realPointX = pt[0] + int(templateArray['width']*0.5)
        realPointY = pt[1] + int(templateArray['height'])
        newRealPoint = (realPointX,realPointY)

        if(len(points) >= 1):
            isNewPoint = checkPoints(newRealPoint, points, templateArray['width'], templateArray['height'])
            if (isNewPoint == True):
                points.append(newRealPoint)
        else:
            points.append(newRealPoint)

    pywinauto.mouse.click(coords=(points[0]))
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


noxWindowObject = getWindowObject(APP_PATH)
bringAppToFront(noxWindowObject)
noxWindowDimensions = getWindowDimensions(noxWindowObject)

matched = initialMatch()

if(isKizunaSP4 == True and len(matched) >= 5):
    for x in range(0, 5):
        matched = chooseEnemy(matched)
    switchFleet()
    chooseBoss(noxWindowDimensions)
else:
    print ('Error less than 5 boats selected')

print('End of script, following array should be left over enemies:')
print(matched)
