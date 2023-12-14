from win32api import GetCursorPos
from keyboard import is_pressed
from PyQt5 import QtWidgets
from PyQt5.QtGui import *
from PIL import ImageGrab
from UI import Ui_Form
import configparser
import numpy as np
import threading
import win32gui
import time
import cv2
import sys

config = configparser.ConfigParser()
inipath = "./config.ini"
config.read(inipath, encoding="utf-8-sig")
key = config["user"]["key"]
path = config["user"]["path"]


class MyWindow(QtWidgets.QWidget, Ui_Form):
    def __init__(self, parent=None):
        super(MyWindow, self).__init__(parent)
        self.setupUi(self)
        self.label.setText(path)
        self.setUpdatesEnabled(True)
        self.ui()
        self.ocv = True
        self.img_black = QImage(260, 260, QImage.Format_RGB888)
        self.img_black.fill(qRgb(0, 0, 0))
        self.img_white = QImage(260, 260, QImage.Format_RGB888)
        self.img_white.fill(qRgb(255, 255, 255))

    def ui(self):
        self.label_cv = QtWidgets.QLabel(self)
        self.label_cv.setGeometry(0, 50, 260, 260)

    def closeEvent(self):
        self.ocv = False

    def opencv(self):
        while self.ocv:
            MS_hwnd = win32gui.FindWindow(None, "MapleStory")
            fig = win32gui.GetWindowRect(MS_hwnd)
            CursorPos = GetCursorPos()
            fig = (CursorPos[0], fig[1] + 32, CursorPos[0] + 260, fig[3] - 8)
            img = ImageGrab.grab(bbox=fig)
            frame = np.array(img)
            try:
                for i in range(CursorPos[1], len(frame)):
                    if (frame[i][259] != [238, 238, 238]).all():
                        break
                    j = i
                for i in range(CursorPos[1], 0, -1):
                    if (frame[i][259] != [238, 238, 238]).all():
                        break
                    k = i
                frame = frame[k - 8 if k > 8 else 0 : j + 8]
            except:
                pass
            if is_pressed(key):
                RGB_f = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                name = time.strftime("%Y%m%d-%H-%M-%S", time.localtime())
                name = path + "/" + str(name) + ".png"
                # cv2.imwrite(name, RGB_f)
                cv2.imencode(".png", RGB_f)[1].tofile(name)
                self.label_cv.setPixmap(QPixmap.fromImage(self.img_white))
                continue
            frame = frame[:261]
            height, width, channel = frame.shape
            bytesPerline = channel * width
            if (frame[10][259] == [238, 238, 238]).all():
                qimg = QImage(frame, width, height, bytesPerline, QImage.Format_RGB888)
                self.label_cv.setPixmap(QPixmap.fromImage(qimg))
            else:
                self.label_cv.setPixmap(QPixmap.fromImage(self.img_black))

    def choose_path(self):
        global path
        Path = QtWidgets.QFileDialog.getExistingDirectory()
        Path = "".join(Path)
        self.label.setText(Path)
        config["user"]["path"] = self.label.text()
        path = Path
        with open(inipath, "w", encoding="utf-8-sig") as configfile:
            config.write(configfile)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    Form = MyWindow()
    Form.setWindowTitle("MSCapture")
    Preview = threading.Thread(target=Form.opencv)
    Preview.start()
    Form.show()
    sys.exit(app.exec_())
