from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
import sys
import win32com.client


speak = win32com.client.Dispatch("SAPI.SpVoice")

ui,_ = loadUiType("GUI.ui")

class MainApp(QMainWindow,ui):
    def __init__(self,parent=None):
        super(MainApp,self).__init__(parent)
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.setWindowTitle("Text to speech")
        self.actionButton()

    def actionButton(self):
        self.pushButton.clicked.connect(self.speakText)
        self.pushButton_1.clicked.connect(self.clearText)
        self.pushButton_2.clicked.connect(self.quitP)

    def speakText(self):
        text = self.lineEdit.text()
        if text == "" or text == " ":
            speak.speak("text box is empty")
            speak.runAndWait()

        else:
            speak.speak(text)
            speak.runAndWait()

    def clearText(self):
        self.lineEdit.clear()

    def quitP(self):
        speak.speak("Thankyou for using this Program")
        speak.runAndWait()
        exit()

def main():
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    app.exec_()
    runAndWait()

if __name__ == "__main__":
    main()
