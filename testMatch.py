import pywinauto
import win32gui
import cv2
import numpy as np
import time
import random

from matplotlib import pyplot as plt
from mss import mss
from PIL import Image

APP_PATH = "D:\\Program Files\\Nox\\bin\\Nox.exe"
UI_WIDTH_720 = 1280
SHOW_MATCH = True

class Vision:
    def __init__(self):
        self.static_templates = {
            'T4-Box'      : 'assets/T4-Box.png',
            'Gear'        : 'assets/Gear.png',
            'Akagi'       : 'assets/Akagi.png',
            'Akashi'      : 'assets/Akashi.png',
            'Aircraft1'   : 'assets/Ac1.png',
            'Battleship1' : 'assets/Bs1.png',
            'Kizuna1'     : 'assets/KZ1.png',
            'Kizuna2'     : 'assets/KZ2.png',
            'Kizuna3'     : 'assets/KZ3.png'
        }
        self.templates = { k: cv2.imread(v, 0) for (k, v) in self.static_templates.items() }
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

    def scaled_find_template(self, name, threshold, scales, image=None):
        if image is None:
            if self.frame is None:
                self.refresh_frame()

            image = self.frame

        initial_template = self.templates[name]
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

    # def match_template2(self, img_grayscale, template, threshold):
    #     """
    #     Matches template image in a target grayscaled image
    #     """
    #
    #     # res = cv2.matchTemplate(img_grayscale, template, cv2.TM_CCOEFF_NORMED)
    #     # matches = np.where(res >= threshold)
    #     # return matches
    #
    #     # Initiate ORB detector
    #     orb = cv2.ORB_create()
    #     # find the keypoints and descriptors with ORB
    #     kp1, des1 = orb.detectAndCompute(template,None)
    #     kp2, des2 = orb.detectAndCompute(img_grayscale,None)
    #     # create BFMatcher object
    #     bf = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
    #     # Match descriptors.
    #     matches = bf.match(des1,des2) #dmmatch
    #     # Sort them in the order of their distance.
    #     matches = sorted(matches, key = lambda x:x.distance)
    #     # Draw first 10 matches.
    #     img3 = cv2.drawMatches(template, kp1, img_grayscale, kp2,matches[:10], None, flags=2)
    #     plt.imshow(img3),plt.show()

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

def checkPoints(newPoint, w, h):
    for point in matched:
        if (abs(newPoint[0] - point[0]) > w or abs(newPoint[1] - point[1]) > h):
            isNewPoint = True
        else:
            isNewPoint = False
    return isNewPoint

noxWindowObject = getWindowObject(APP_PATH)
bringAppToFront(noxWindowObject)
noxWindowDimensions = getWindowDimensions(noxWindowObject)

vision = Vision()
vision.setMonitor(noxWindowDimensions)

vision.save_screenshot()

UIscale = vision.find_scale(noxWindowDimensions['width'])
scales = [UIscale]
matched = []

initial_template = vision.templates['Kizuna1']
scaled_template = cv2.resize(initial_template, (0,0), fx=UIscale, fy=UIscale)
w, h = scaled_template.shape[::-1]
# matches = vision.scaled_find_template('Akagi', 0.5, scales=scales)
# matches = vision.find_template('Gear', 0.9)

# TODO only works in top left corner now
matches = vision.scaled_find_template('Kizuna1', 0.65, scales=scales)
# print (np.shape(matches)[1] + noxWindowDimensions['left'])
# print (np.shape(matches)[0] + noxWindowDimensions['top'])

for pt in zip(*matches[::-1]):
    # add the template size to point
    realPointX = pt[0] + int(w*0.5)
    realPointY = pt[1] + int(h*0.5)
    newRealPoint = (realPointX,realPointY)

    if(len(matched) >= 1):
        isNewPoint = checkPoints(newRealPoint, w, h)
        if (isNewPoint == True):
            matched.append(newRealPoint)
    else:
        matched.append(newRealPoint)

initial_template = vision.templates['Kizuna2']
scaled_template = cv2.resize(initial_template, (0,0), fx=UIscale, fy=UIscale)
w, h = scaled_template.shape[::-1]
# matches = vision.scaled_find_template('Akagi', 0.5, scales=scales)
# matches = vision.find_template('Gear', 0.9)

# TODO only works in top left corner now
matches = vision.scaled_find_template('Kizuna2', 0.65, scales=scales)
# print (np.shape(matches)[1] + noxWindowDimensions['left'])
# print (np.shape(matches)[0] + noxWindowDimensions['top'])

for pt in zip(*matches[::-1]):
    # add the template size to point
    realPointX = pt[0] + int(w*0.5)
    realPointY = pt[1] + int(h*0.5)
    newRealPoint = (realPointX,realPointY)

    if(len(matched) >= 1):
        isNewPoint = checkPoints(newRealPoint, w, h)
        if (isNewPoint == True):
            matched.append(newRealPoint)
    else:
        matched.append(newRealPoint)

initial_template = vision.templates['Kizuna3']
scaled_template = cv2.resize(initial_template, (0,0), fx=UIscale, fy=UIscale)
w, h = scaled_template.shape[::-1]
# matches = vision.scaled_find_template('Akagi', 0.5, scales=scales)
# matches = vision.find_template('Gear', 0.9)

# TODO only works in top left corner now
matches = vision.scaled_find_template('Kizuna3', 0.65, scales=scales)
# print (np.shape(matches)[1] + noxWindowDimensions['left'])
# print (np.shape(matches)[0] + noxWindowDimensions['top'])

for pt in zip(*matches[::-1]):
    # add the template size to point
    realPointX = pt[0] + int(w*0.5)
    realPointY = pt[1] + int(h*0.5)
    newRealPoint = (realPointX,realPointY)

    if(len(matched) >= 1):
        isNewPoint = checkPoints(newRealPoint, w, h)
        if (isNewPoint == True):
            matched.append(newRealPoint)
    else:
        matched.append(newRealPoint)

print(matched)

randNumb = random.randint(1, len(matched))
print (matched[randNumb-1][0],matched[randNumb-1][1])
#pywinauto.mouse.click(coords=(matched[randNumb][0],matched[randNumb][1]))

#print (np.shape(matches)[1] >= 1)
