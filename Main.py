import pywinauto

APP_PATH = "D:\\Program Files\\Nox\\bin\\Nox.exe"

def getWindowObject(appPath):
    app = pywinauto.Application().connect(path=appPath)
    window = app.top_window()
    return window

def bringAppToFront(appWindow):
    #bring window into foreground
    if appWindow.has_style(pywinauto.win32defines.WS_MINIMIZE): # if minimized
        pywinauto.win32functions.ShowWindow(appWindow.wrapper_object(), 9) # restore window state
    else:
        pywinauto.win32functions.SetForegroundWindow(appWindow.wrapper_object()) #bring to front

noxWindowObject = getWindowObject(APP_PATH)
bringAppToFront(noxWindowObject)
