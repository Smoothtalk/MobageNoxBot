import pywinauto
import win32gui

APP_PATH = "D:\\Program Files\\Nox\\bin\\Nox.exe"

def getWindowObject(appPath):
    app = pywinauto.Application().connect(path=appPath)
    #window = app.top_window()
    window = app.window(title_re="NoxPlayer.*")
    return window

def bringAppToFront(appWindow):
    #bring window into foreground
    if appWindow.has_style(pywinauto.win32defines.WS_MINIMIZE): # if minimized
        pywinauto.win32functions.ShowWindow(appWindow.wrapper_object(), 9) # restore window state
    else:
        pywinauto.win32functions.SetForegroundWindow(appWindow.wrapper_object()) #bring to front

def getWindowDimensions(appWindow):
    print (type(appWindow.rectangle()))
    w = appWindow.rectangle().width()
    h = appWindow.rectangle().height() - 30 #window element bar subtraction
    x = appWindow.rectangle().left
    y = appWindow.rectangle().top

    print ("top left co-ords:" + str(x) + 'x' + str(y))
    print ("size:" + str(w) + 'x' + str(h))

noxWindowObject = getWindowObject(APP_PATH)
bringAppToFront(noxWindowObject)
getWindowDimensions(noxWindowObject)
