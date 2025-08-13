import configparser
import sys
import threading
import time

import cv2
import numpy as np
import win32gui
from keyboard import is_pressed
from PIL import ImageGrab
from PyQt6 import QtWidgets
from PyQt6.QtGui import *
from win32api import GetCursorPos

from UI_files.UI import Ui_Form

config = configparser.ConfigParser()
inipath = "./config.ini"
config.read(inipath, encoding="utf-8-sig")
key = config["user"]["key"]
path = config["user"]["path"]

BORDER_COLOR = np.array([238, 238, 238])
BLACK_COLOR = (0, 0, 0)
WHITE_COLOR = (255, 255, 255)


class MyWindow(QtWidgets.QWidget, Ui_Form):
    def __init__(self, parent=None):
        super(MyWindow, self).__init__(parent)
        self.setupUi(self)
        self.label.setText(path)
        self.setUpdatesEnabled(True)
        self.ui()
        self.ocv = True

        self.img_black = QImage(260, 260, QImage.Format.Format_RGB888)
        self.img_black.fill(qRgb(*BLACK_COLOR))
        self.img_white = QImage(260, 260, QImage.Format.Format_RGB888)
        self.img_white.fill(qRgb(*WHITE_COLOR))

        self.MS_hwnd = None
        self.last_hwnd_check = 0

    def ui(self):
        self.label_cv = QtWidgets.QLabel(self)
        self.label_cv.setGeometry(0, 50, 260, 260)

    def closeEvent(self):
        self.ocv = False

    def get_maple_window(self):
        current_time = time.time()
        if self.MS_hwnd is None or current_time - self.last_hwnd_check > 5:
            MS_hwnd1 = win32gui.FindWindow(None, "MapleStory")
            MS_hwnd2 = win32gui.FindWindowEx(
                None, MS_hwnd1, None, "MapleStory")
            self.MS_hwnd = MS_hwnd1 if MS_hwnd2 == 0 else MS_hwnd2
            self.last_hwnd_check = current_time
        return self.MS_hwnd

    def find_border_optimized(self, frame, cursor_y):
        height = len(frame)
        if height == 0 or cursor_y >= height:
            return 0, height

        right_column = frame[:, 259]
        is_border = np.all(right_column == BORDER_COLOR, axis=1)

        non_border_indices = np.where(~is_border)[0]

        if len(non_border_indices) == 0:
            return cursor_y, cursor_y

        upper_indices = non_border_indices[non_border_indices < cursor_y]
        lower_indices = non_border_indices[non_border_indices > cursor_y]

        k = upper_indices[-1] if len(upper_indices) > 0 else 0
        j = lower_indices[0] if len(lower_indices) > 0 else height - 1

        return max(k - 6, 0), min(j + 8, height)

    def opencv(self):
        last_capture_time = 0
        capture_cooldown = 0.1

        while self.ocv:
            try:
                MS_hwnd = self.get_maple_window()
                if not MS_hwnd:
                    time.sleep(0.1)
                    continue

                fig = win32gui.GetWindowRect(MS_hwnd)
                CursorPos = GetCursorPos()

                bbox = (CursorPos[0], fig[1] + 32,
                        CursorPos[0] + 260, fig[3] - 8)

                img = ImageGrab.grab(bbox=bbox, all_screens=False)
                frame = np.asarray(img, dtype=np.uint8)  # 使用asarray避免複製

                relative_cursor_y = CursorPos[1] - (fig[1] + 32)

                if relative_cursor_y >= 0 and relative_cursor_y < len(frame):
                    k, j = self.find_border_optimized(frame, relative_cursor_y)
                    frame = frame[k:j]

                if is_pressed(key):
                    current_time = time.time()
                    if current_time - last_capture_time > capture_cooldown:
                        frame_to_save = frame[:,
                                              1:] if frame.shape[1] > 1 else frame
                        RGB_f = cv2.cvtColor(frame_to_save, cv2.COLOR_BGR2RGB)
                        name = f"{path}/{time.strftime('%Y%m%d-%H-%M-%S', time.localtime())}.png"
                        cv2.imencode(".png", RGB_f)[1].tofile(name)
                        self.label_cv.setPixmap(
                            QPixmap.fromImage(self.img_white))
                        last_capture_time = current_time
                    continue

                if len(frame) > 261:
                    frame = frame[:261]

                if len(frame) > 10 and frame.shape[1] >= 260:
                    if np.all(frame[10, 259] == BORDER_COLOR):
                        height, width, channel = frame.shape
                        bytesPerline = channel * width
                        qimg = QImage(frame.data, width, height,
                                      bytesPerline, QImage.Format.Format_RGB888)
                        self.label_cv.setPixmap(QPixmap.fromImage(qimg))
                    else:
                        self.label_cv.setPixmap(
                            QPixmap.fromImage(self.img_black))
                else:
                    self.label_cv.setPixmap(QPixmap.fromImage(self.img_black))

                time.sleep(0.033)

            except Exception as e:
                # print(f"Error in opencv loop: {e}")
                time.sleep(0.1)

    def choose_path(self):
        global path
        Path = QtWidgets.QFileDialog.getExistingDirectory()
        if Path:
            self.label.setText(Path)
            config["user"]["path"] = Path
            path = Path
            with open(inipath, "w", encoding="utf-8-sig") as configfile:
                config.write(configfile)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    Form = MyWindow()
    Form.setWindowTitle("MSCapture")
    Preview = threading.Thread(
        target=Form.opencv, daemon=True)
    Preview.start()
    Form.show()
    sys.exit(app.exec())
