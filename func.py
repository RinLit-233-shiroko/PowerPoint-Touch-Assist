import pyautogui
import pygetwindow as gw
from pynput import mouse
from win32 import win32api
import conf_file as conf

#get分辨率
screenX = win32api.GetSystemMetrics(0)
screenY = win32api.GetSystemMetrics(1)

#缩放比率
scale_rate = conf.dpi_dict[conf.read_conf('General', 'DPI')]

ppt_window_title = conf.read_conf('General', 'PPT_Title')

tsk_edge = screenY - 150*scale_rate
ppt_menuWL = 125*scale_rate
ppt_menuWR = screenX - 125*scale_rate

is_pressed = False
pressed_menuButton = False
pos_mouse = [0,0]

#检测是否有PowerPoint放映
def is_powerpoint_showing():
    try:
        windows = gw.getWindowsWithTitle(ppt_window_title)
        return len(windows) > 0
    except Exception:
        print(f'这是\n{Exception}\n喵')#QwQ
    return False

def is_finger_not_slide(x, y):
    if [x, y] == pos_mouse:
        return True
    else:
        return False
    
def is_not_click_taskbar(mouse_y):
    if mouse_y <= tsk_edge:
        return True
    else:
        return False
    
def is_not_click_ppt_menubar(mouse_x):
    if mouse_x >= ppt_menuWL and mouse_x <= ppt_menuWR:
        return True
    else:
        return False

#点击事件
def on_click(x, y, button, pressed):
    global is_pressed
    global pos_mouse
    global pressed_menuButton
    if button == mouse.Button.right:
        pressed_menuButton = True
    if button == mouse.Button.left:
        if pressed:
            is_pressed = True
            pos_mouse = [x, y]
            ###如果点击ppt工具栏，禁用下一次左键点击以操作弹出的菜单###
            if is_not_click_ppt_menubar(x) is False or is_not_click_taskbar(y) is False:
                pressed_menuButton = True
            else:
                pressed_menuButton = False
        else:#当放开时判定是否按下且符合条件
            if is_pressed and pressed_menuButton == False and is_finger_not_slide(x, y) and is_powerpoint_showing():
                if is_not_click_ppt_menubar(x) and is_not_click_taskbar(y):
                    pyautogui.press('space')
            is_pressed = False
            

#监听
def main():
    print('程序运行中……\nRinLit_233OuO @bilibili')
    with mouse.Listener(on_click=on_click) as listener:
        listener.join()

if __name__ == '__main__':
    main()  