'''
    PowerPoint Touch Assist v1.1
    设置菜单
    Author: RinLit_233OuO @bilibili
    Version:  1.1
'''

from PyQt6.QtWidgets import QApplication, QDialog, QLineEdit, QPushButton, QComboBox, QCheckBox
from PyQt6 import uic
from PyQt6.QtCore import QUrl, Qt
from PyQt6.QtGui import QDesktopServices
import sys
import conf_file as config
import shortcut as a

class FramelessWindow(QDialog):
    def __init__(self):
        super().__init__()
        uic.loadUi('./settings.ui', self)

        ###控件###
        #选项1 [DPI]
        self.opt1_Combo = self.findChild(QComboBox, 'opt1_Combo')
        self.opt1_Combo.setCurrentIndex(int(config.read_conf('General', 'DPI')))
        self.opt1_Combo.currentIndexChanged.connect(self.opt1_Save)
        #选项2 [ppt标题]
        self.opt2_LineEdit = self.findChild(QLineEdit, 'opt2_LineEdit')
        self.opt2_LineEdit.setText(config.read_conf('General', 'PPT_Title'))
        self.opt2_LineEdit.textChanged.connect(self.opt2_Save)
        #选项3 [开机自启动]
        self.opt3_checkBox = self.findChild(QCheckBox, 'opt3_checkBox')
        self.opt3_checkBox.setChecked(int(config.read_conf('General', 'auto_startup')))
        self.opt3_checkBox.stateChanged.connect(self.opt3_Save)
        #reset按钮
        self.opt2_Reset = self.findChild(QPushButton, 'opt2Reset')
        self.opt2_Reset.clicked.connect(self.opt2_ResetToDefault)
        #关闭按钮
        self.close_button = self.findChild(QPushButton, 'Close')
        self.close_button.clicked.connect(self.close_window)
        #b站启动按钮
        self.bilibili_button = self.findChild(QPushButton, 'pushButton_Bilibili')
        self.bilibili_button.clicked.connect(self.RedirectTo_Bilibili)
        #关于按钮
        self.about_button = self.findChild(QPushButton, 'pushButton_about')
        self.about_button.clicked.connect(self.RedirectTo_About)
        #Github
        self.github_button = self.findChild(QPushButton, 'pushButton_Github')
        self.github_button.clicked.connect(self.RedirectTo_Github)

    #behavior行为
    def opt1_Save(self, text):
        config.write_conf('General', 'DPI', str(text))  
    def opt2_Save(self, text):
        config.write_conf('General', 'PPT_Title', text)
    def opt2_ResetToDefault(self):
        config.write_conf('General', 'PPT_Title', 'PowerPoint 幻灯片放映')
        self.opt2_LineEdit.setText(config.read_conf('General', 'PPT_Title'))
    def opt3_Save(self, state): 
        is_checked = state == Qt.CheckState.Checked.value
        config.write_conf('General', 'auto_startup', str(int(is_checked)))
        if is_checked:
            a.add_to_startup('PowerPoint_TouchAssist.exe')
        else:
            a.remove_from_startup()

    def close_window(self):
        sys.exit()
    def RedirectTo_About(self):
        url = QUrl('https://pptfortouch.framer.website/')
        QDesktopServices.openUrl(url)
    def RedirectTo_Bilibili(self):
        url = QUrl('https://space.bilibili.com/569522843')
        QDesktopServices.openUrl(url)
    def RedirectTo_Github(self):#gayhub
        url = QUrl('https://github.com/RinLit-233-shiroko/PowerPoint-Touch-Assist')
        QDesktopServices.openUrl(url)
    


def main():
    app = QApplication(sys.argv)
    window = FramelessWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()