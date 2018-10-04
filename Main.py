import pywinauto
import win32gui
import cv2
import numpy as np
import time

from matplotlib import pyplot as plt
from mss import mss
from PIL import Image

APP_PATH = "D:\\Program Files\\Nox\\bin\\Nox.exe"
UI_WIDTH_720 = 1280

class Vision:
    def __init__(self):
        self.static_templates = {
            'T4-Box' : 'assets/T4-Box.png',
            'Gear' :   'assets/Gear.png'
        }
        self.templates = { k: cv2.imread(v, 0) for (k, v) in self.static_templates.items() }
        self.screen = mss()
        self.frame = None

    def setMonitor(self, monitorDimensions):
        self.monitor = {'top': monitorDimensions['top'], 'left': monitorDimensions['left'], 'width': monitorDimensions['width'], 'height': monitorDimensions['height']}

    def convert_rgb_to_bgr(self, img):
        return img[:, :, ::-1]

    def take_screenshot(self):
        print (self.monitor)
        sct_img = self.screen.grab(self.monitor)
        img = Image.frombytes('RGB', sct_img.size, sct_img.rgb)
        img = np.array(img)
        img = self.convert_rgb_to_bgr(img)
        img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        return img_gray

    def refresh_frame(self):
        self.frame = self.take_screenshot()

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

    def find_scale_template(self, currentWidth):
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

noxWindowObject = getWindowObject(APP_PATH)
#bringAppToFront(noxWindowObject)
noxWindowDimensions = getWindowDimensions(noxWindowObject)

vision = Vision()
vision.setMonitor(noxWindowDimensions)

UIscale = vision.find_scale_template(noxWindowDimensions['width'])
scales = [1.5, 1.0, UIscale]

matches = vision.scaled_find_template('Gear', 0.8, scales=scales)
# matches = vision.find_template('Gear', 0.9)

print (np.shape(matches)[1] >= 1)
