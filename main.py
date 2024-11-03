"""
    PowerPoint Touch Assist v1.2

    Author: RinLit_233OuO @bilibili
    Version:  1.2

    ###
    v1.2: 修复bug
    v1.1: 增加高分辨率支持, 增加设置菜单
    v1: 新App出现!
"""
import sys
import threading

from PyQt6 import uic
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QCursor
from PyQt6.QtWidgets import QApplication, QWidget

import conf_file
import conf_ui
import func
import shortcut as s


class FramelessWindow(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi('./main.ui', self)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.m_flag = False
        QTimer.singleShot(3000, self.hide)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.m_flag = True
            self.m_Position = event.globalPosition().toPoint() - self.pos()  # 获取鼠标相对窗口的位置
            event.accept()

    def mouseMoveEvent(self, event):
        if Qt.MouseButton.LeftButton and self.m_flag:
            self.move(event.globalPosition().toPoint() - self.m_Position)  # 更改窗口位置
            event.accept()

    def mouseReleaseEvent(self, event):
        self.m_flag = False
        self.setCursor(QCursor(Qt.CursorShape.ArrowCursor))


def run_func():
    func.main()


def run_settings():
    conf_ui.main()


def main():
    args = sys.argv[1:]
    # 如果存在参数"settings"就运行设置
    if "settings" in args:
        run_settings()
    # 是否初次启动，是打开设置，不是 直接_ _启动
    if bool(conf_file.read_conf('Miscellaneous', 'InitialStartUp')):
        conf_file.write_conf('Miscellaneous', 'InitialStartUp', '0')
        s.add_to_desktop('PowerPoint_TouchAssist.exe')
        s.add_to_menu('PowerPoint_TouchAssist.exe')
        s.add_to_menu('PowerPoint_TouchAssist.exe', name='PowerPoint 触屏辅助 - 设置', args='settings')
        run_settings()
    else:
        # 启动func.py中的main
        func_thread = threading.Thread(target=run_func)
        func_thread.daemon = True
        func_thread.start()

    # _ _,启动!
    app = QApplication(sys.argv)
    window = FramelessWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
