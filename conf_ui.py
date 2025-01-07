import sys

from PyQt6 import uic
from PyQt6.QtCore import QUrl, Qt
from PyQt6.QtGui import QDesktopServices, QIcon
from PyQt6.QtWidgets import QApplication, QScroller
from qfluentwidgets import FluentWindow, FluentIcon as fIcon, setTheme, Theme, ComboBox, setThemeColor, LineEdit, \
    ToolButton, SwitchButton, SmoothScrollArea, NavigationItemPosition, BodyLabel, PushButton

import config as config
import shortcut as a


def open_about_url():
    url = QUrl('https://pptfortouch.framer.website/')
    QDesktopServices.openUrl(url)


def open_bilibili():
    url = QUrl('https://space.bilibili.com/569522843')
    QDesktopServices.openUrl(url)


def open_github():  # GayHub
    url = QUrl('https://github.com/RinLit-233-shiroko/PowerPoint-Touch-Assist')
    QDesktopServices.openUrl(url)


def save_if_auto_start(state):
    is_checked = state == Qt.CheckState.Checked.value
    config.write_conf('General', 'auto_startup', str(1 if state else 0))
    if is_checked:
        a.add_to_startup('PowerPoint_TouchAssist.exe')
    else:
        a.remove_from_startup()


# def save_if_use_regex(state):
#     is_checked = state == Qt.CheckState.Checked.value
#     config.write_conf('General', 'use_regex', str(is_checked))


def save_dpi(text):
    config.write_conf('General', 'DPI', config.dpi_list[text])


class Settings(FluentWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.initUI()

    def initUI(self):
        setThemeColor('#e66d4a')
        setTheme(Theme.AUTO)

        self.init_nav()
        self.add_subInterface()
        self.setupPages()

        self.setWindowTitle("PPT 触屏辅助 | 设置")
        self.setWindowIcon(QIcon("img/favicon.ico"))
        self.setMinimumSize(500, 400)

        screen_geometry = QApplication.primaryScreen().geometry()
        screen_width = screen_geometry.width()
        screen_height = screen_geometry.height()

        width = int(screen_width * 0.4)
        height = int(screen_height * 0.5)

        self.move(int(screen_width / 2 - width / 2), int(screen_height / 2 - height / 2))
        self.resize(width, height)

    def init_nav(self):
        self.settingsPage = uic.loadUi("ui/settings.ui")
        self.settingsPage.setObjectName("settings")
        self.aboutPage = uic.loadUi("ui/about.ui")
        self.aboutPage.setObjectName("about")

    def add_subInterface(self):
        self.navigationInterface.setExpandWidth(175)
        self.addSubInterface(self.settingsPage, fIcon.SETTING, "设置")
        self.addSubInterface(self.aboutPage, fIcon.INFO, "关于", NavigationItemPosition.BOTTOM)

    def setupPages(self):
        self.setup_settings()
        self.setup_about()

    def setup_settings(self):
        general_scroll = self.settingsPage.findChild(SmoothScrollArea, 'general_scroll')
        QScroller.grabGesture(general_scroll.viewport(), QScroller.ScrollerGestureType.LeftMouseButtonGesture)

        self.dpi_select = self.settingsPage.findChild(ComboBox, 'dpi_select')
        self.program_name = self.settingsPage.findChild(LineEdit, 'program_name')
        self.reset_name = self.settingsPage.findChild(ToolButton, 'reset_name')
        self.switch_startup = self.settingsPage.findChild(SwitchButton, 'switch_startup')

        self.dpi_select.addItems(config.dpi_list_text)
        self.dpi_select.setCurrentIndex(config.dpi_list.index(config.read_conf('General', 'DPI')))
        self.dpi_select.currentIndexChanged.connect(save_dpi)
        self.program_name.setFixedWidth(225)
        self.program_name.setText(config.read_conf('General', 'program_title'))
        self.program_name.textChanged.connect(lambda text: config.write_conf('General', 'program_title', text))
        self.reset_name.setIcon(fIcon.CANCEL)
        self.reset_name.clicked.connect(self.reset_ppt_title)
        self.switch_startup.setChecked(config.read_conf('General', 'auto_startup') == '1')
        self.switch_startup.checkedChanged.connect(save_if_auto_start)

    def setup_about(self):
        about_scroll = self.aboutPage.findChild(SmoothScrollArea, 'about_scroll')
        QScroller.grabGesture(about_scroll.viewport(), QScroller.ScrollerGestureType.LeftMouseButtonGesture)

        version = self.aboutPage.findChild(BodyLabel, 'version')
        btn_github = self.aboutPage.findChild(PushButton, 'button_github')
        btn_bilibili = self.aboutPage.findChild(PushButton, 'button_bilibili')
        btn_website = self.aboutPage.findChild(PushButton, 'button_website')

        btn_github.clicked.connect(open_github)
        btn_bilibili.clicked.connect(open_bilibili)
        btn_website.clicked.connect(open_about_url)
        version.setText(f"版本：{config.read_conf('Misc', 'ver')}")

    # Methods
    def reset_ppt_title(self):
        config.write_conf('General', 'program_title', 'PowerPoint 幻灯片放映')
        self.program_name.setText(config.read_conf('General', 'program_title'))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Settings()
    window.show()
    sys.exit(app.exec())

