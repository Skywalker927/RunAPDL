import win32gui
import win32con
import time
import os
import pickle
import sys

# 获取输入参数
paras = sys.argv

# 文件目录
path = os.path.abspath('.')

# 复制文件到默认位置
file1 = paras[1]
directory, name = os.path.split(paras[1])
_Fname, _Ext = os.path.splitext(name)
_Ext = _Ext[1:]
_input = "".join(["/INPUT,", _Fname, ",", _Ext,",",directory, "\\"])

# 寻找窗口
hwndname = os.path.join(path,'hwnd.data')
indent1 = 'Ansys Mechanical Enterprise Utility Menu'
# 注意！！！！！需将indent1变量改成ansys命令流软件的窗口名称
indent2 = 'Output Window'

try:
    with open(hwndname,'rb') as f:
        hwnd = pickle.load(f)
        hwnd1 = hwnd[0]
        hwnd2 = hwnd[1]
    title = win32gui.GetWindowText(hwnd1)
    if indent1  not in title:
        raise ValueError('提示：未找到已开启的Ansys程序 或 需添加新的窗口名')

except:
    hwnd_title = dict()
    def get_all_hwnd(hwnd,mouse):
        if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
            hwnd_title.update({hwnd:win32gui.GetWindowText(hwnd)})
    
    win32gui.EnumWindows(get_all_hwnd, 0)
    for h,t in hwnd_title.items():
        if indent1 in t:
            hwnd1 = h
            break
    for h,t in hwnd_title.items():
        if indent2 in t:
            hwnd2 = h
            break
    
    hwnd = [hwnd1,hwnd2]
    with open(hwndname,'wb') as f:
        pickle.dump(hwnd,f)

title = win32gui.GetWindowText(hwnd1)
if indent1  not in title:
    print('---------------------------------------------------------')
    print('提示：未找到已开启的Ansys程序 或 需添加新的窗口名')
    print('---------------------------------------------------------')

try:
    win32gui.ShowWindow(hwnd1, win32con.SW_MAXIMIZE)
    win32gui.ShowWindow(hwnd1,win32con.SW_SHOW)
    time.sleep(0.1)
    win32gui.ShowWindow(hwnd2, win32con.SW_MINIMIZE)
    temp = win32gui.SetForegroundWindow(hwnd1)
except:
    pass
finally:
    time.sleep(0.1)
    WM_CHAR = 0x0102
    for char in _input:
        win32gui.SendMessage(hwnd1, WM_CHAR, ord(char), None)
    time.sleep(0.1)    
    win32gui.SendMessage(hwnd1, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    win32gui.SendMessage(hwnd1, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

os.remove(hwndname)