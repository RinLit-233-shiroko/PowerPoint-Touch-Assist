import os

from win32com.client import Dispatch

program_name = 'PowerPoint 触屏辅助'


def create_shortcut(desktop_folder, file_path, icon_path):
    shortcut_path = os.path.join(desktop_folder, f'{program_name}.lnk')

    # 创建快捷方式
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.TargetPath = file_path
    shortcut.WorkingDirectory = os.path.dirname(file_path)
    shortcut.IconLocation = icon_path  # 设置图标路径
    shortcut.save()


def add_to_startup(file_path="", icon_path=""):
    if file_path == "":
        file_path = os.path.realpath(__file__)
    else:
        file_path = os.path.abspath(file_path)  # 将相对路径转换为绝对路径

    if icon_path == "":
        icon_path = file_path  # 如果未指定图标路径，则使用程序路径

    # 获取启动文件夹路径
    startup_folder = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
    # 创建快捷方式
    create_shortcut(startup_folder, file_path, icon_path)


def remove_from_startup():
    startup_folder = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
    shortcut_path = os.path.join(startup_folder, f'{program_name}.lnk')
    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)


def add_to_desktop(file_path="", icon_path=""):
    if file_path == "":
        file_path = os.path.realpath(__file__)
    else:
        file_path = os.path.abspath(file_path)  # 将相对路径转换为绝对路径

    if icon_path == "":
        icon_path = file_path  # 如果未指定图标路径，则使用程序路径
    # 获取文件夹路径
    desktop_folder = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    # 创建快捷方式
    create_shortcut(desktop_folder, file_path, icon_path)


def add_to_menu(file_path="", icon_path="", name='PowerPoint 触屏辅助', args=''):
    if file_path == "":
        file_path = os.path.realpath(__file__)
    else:
        file_path = os.path.abspath(file_path)  # 将相对路径转换为绝对路径

    if icon_path == "":
        icon_path = file_path  # 如果未指定图标路径，则使用程序路径
    # 获取文件夹路径
    menu_folder = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs')
    # 快捷方式文件名
    shortcut_path = os.path.join(menu_folder, f'{name}.lnk')

    # 创建快捷方式
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.TargetPath = file_path
    shortcut.Arguments = args
    shortcut.WorkingDirectory = os.path.dirname(file_path)
    shortcut.IconLocation = icon_path  # 设置图标路径
    shortcut.save()


if __name__ == '__main__':
    add_to_startup('PowerPoint_TouchAssist.exe')
