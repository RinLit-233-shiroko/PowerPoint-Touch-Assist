"""
    PowerPoint Touch Assist
    设置菜单
    Author: RinLit_233OuO @bilibili
    Version:  1.2
"""

import sys

from PyQt6 import uic
from PyQt6.QtCore import QUrl, Qt
from PyQt6.QtGui import QDesktopServices
from PyQt6.QtWidgets import QApplication, QDialog, QLineEdit, QPushButton, QComboBox, QCheckBox

import conf_file as config
import shortcut as a


def close_window():
    sys.exit()


def open_about_url():
    url = QUrl('https://pptfortouch.framer.website/')
    QDesktopServices.openUrl(url)


def open_bilibili():
    url = QUrl('https://space.bilibili.com/569522843')
    QDesktopServices.openUrl(url)


def open_github():  # gayhub
    url = QUrl('https://github.com/RinLit-233-shiroko/PowerPoint-Touch-Assist')
    QDesktopServices.openUrl(url)


def save_if_auto_start(state):
    is_checked = state == Qt.CheckState.Checked.value
    config.write_conf('General', 'auto_startup', str(is_checked))
    if is_checked:
        a.add_to_startup('PowerPoint_TouchAssist.exe')
    else:
        a.remove_from_startup()


def save_if_use_regex(state):
    is_checked = state == Qt.CheckState.Checked.value
    config.write_conf('General', 'use_regex', str(is_checked))


def save_ppt_title(text):
    config.write_conf('General', 'PPT_Title', text)


def save_dpi(text):
    config.write_conf('General', 'DPI', str(text))


class FramelessWindow(QDialog):
    def __init__(self):
        super().__init__()
        uic.loadUi('./settings.ui', self)

        # ## 控件
        # 选项1 [DPI]
        self.opt1_Combo = self.findChild(QComboBox, 'opt1_Combo')
        self.opt1_Combo.setCurrentIndex(int(config.read_conf('General', 'DPI')))
        self.opt1_Combo.currentIndexChanged.connect(save_dpi)
        # 选项2 [ppt标题]
        self.opt2_LineEdit = self.findChild(QLineEdit, 'opt2_LineEdit')
        self.opt2_LineEdit.setText(config.read_conf('General', 'PPT_Title'))
        self.opt2_LineEdit.textChanged.connect(save_ppt_title)
        # 选项3 [开机自启动]
        self.opt3_checkBox = self.findChild(QCheckBox, 'opt3_checkBox')
        self.opt3_checkBox.setChecked(bool(config.read_conf('General', 'auto_startup')))
        self.opt3_checkBox.stateChanged.connect(save_if_auto_start)
        # 选项4 [开机自启动]
        self.opt3_checkBox = self.findChild(QCheckBox, 'opt4_checkBox')
        self.opt3_checkBox.setChecked(bool(config.read_conf('General', 'use_regex')))
        self.opt3_checkBox.stateChanged.connect(save_if_use_regex)
        # reset按钮
        self.opt2_Reset = self.findChild(QPushButton, 'opt2Reset')
        self.opt2_Reset.clicked.connect(self.reset_ppt_title)
        # 关闭按钮
        self.close_button = self.findChild(QPushButton, 'Close')
        self.close_button.clicked.connect(close_window)
        # b站启动按钮
        self.bilibili_button = self.findChild(QPushButton, 'pushButton_Bilibili')
        self.bilibili_button.clicked.connect(open_bilibili)
        # 关于按钮
        self.about_button = self.findChild(QPushButton, 'pushButton_about')
        self.about_button.clicked.connect(open_about_url)
        # Github
        self.github_button = self.findChild(QPushButton, 'pushButton_Github')
        self.github_button.clicked.connect(open_github)

    # behavior行为

    def reset_ppt_title(self):
        config.write_conf('General', 'PPT_Title', 'PowerPoint 幻灯片放映')
        self.opt2_LineEdit.setText(config.read_conf('General', 'PPT_Title'))


def main():
    app = QApplication(sys.argv)
    window = FramelessWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
