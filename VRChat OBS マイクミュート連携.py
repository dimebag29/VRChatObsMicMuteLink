# ==============================================================================================================
# 作成者:dimebag29 作成日:2025年2月13日 バージョン:v0.2
# (Author:dimebag29 Creation date:February 13, 2025 Version:v0.2)
#
# このプログラムのライセンスはLGPLv3です。pystrayライブラリのライセンスを継承しています。
# (This program is licensed to LGPLv3. Inherits the license of the pystray library.)
# https://www.gnu.org/licenses/lgpl-3.0.html
#
# 開発環境 (Development environment)
# ･python 3.7.5
# ･auto-py-to-exe 2.36.0 (used to create the exe file)
#
# exe化時のauto-py-to-exeの設定
# ･ひとつのディレクトリにまとめる (--onedir)
# ･ウィンドウベース (--windowed)
# ･exeアイコン設定 (--icon) (VRChatObsMicMuteLink.ico)
# ･追加ファイルでタスクトレイアイコン追加 (--add-data) (VRChatObsMicMuteLink.ico)
# ==============================================================================================================

# python 3.7.5の標準ライブラリ (Libraries included as standard in python 3.7.5)
import sys
import os
import threading

# 外部ライブラリ (External libraries)
import win32api                                                                 # Included in pywin32 Version:306
import win32con                                                                 # Included in pywin32 Version:306
from pythonosc import dispatcher, osc_server                                    # Version:1.8.3
from PIL import Image                                                           # Version:9.5.0
from pystray import Icon, Menu, MenuItem                                        # Version:0.19.5



def Push_Exit():
    global Icon
    Icon.stop()                                                                 # タスクトレイ常駐終了


def ChangedMuteSelfPara(Address, MuteSelf): 
    # キーコード参考 : https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
    
    # VRChatがミュートなら実行 (キー操作 : Shift + Ctrl + Alt + [)
    if True == MuteSelf:
        win32api.keybd_event(0x10, 0) # Shift
        win32api.keybd_event(0x11, 0) # Ctrl
        win32api.keybd_event(0x12, 0) # Alt
        win32api.keybd_event(0xDB, 0) # [
        win32api.Sleep(100)
        win32api.keybd_event(0xDB, 0, win32con.KEYEVENTF_KEYUP) # [
        win32api.keybd_event(0x12, 0, win32con.KEYEVENTF_KEYUP) # Alt
        win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP) # Ctrl
        win32api.keybd_event(0x10, 0, win32con.KEYEVENTF_KEYUP) # Shift
    
    # VRChatがミュートじゃなかったら実行 (キー操作 : Shift + Ctrl + Alt + ])
    else:
        win32api.keybd_event(0x10, 0) # Shift
        win32api.keybd_event(0x11, 0) # Ctrl
        win32api.keybd_event(0x12, 0) # Alt
        win32api.keybd_event(0xDD, 0) # ]
        win32api.Sleep(100)
        win32api.keybd_event(0xDD, 0, win32con.KEYEVENTF_KEYUP) # ]
        win32api.keybd_event(0x12, 0, win32con.KEYEVENTF_KEYUP) # Alt
        win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP) # Ctrl
        win32api.keybd_event(0x10, 0, win32con.KEYEVENTF_KEYUP) # Shift


def LoggingMuteSelfParaThread():
    dispatcher.map("/avatar/parameters/MuteSelf", ChangedMuteSelfPara)
    server = osc_server.ThreadingOSCUDPServer(("127.0.0.1", 9001), dispatcher)
    server.serve_forever()


# ==============================================================================================================
# OSC関連
dispatcher = dispatcher.Dispatcher()
LoggingMuteSelfParaThreadObj = threading.Thread(target=LoggingMuteSelfParaThread, daemon=True)  # daemon=Trueでデーモン化しないとメインスレッドが終了しても生き残り続けてしまう
LoggingMuteSelfParaThreadObj.start()

# アイコンのパスを取得 https://stackoverflow.com/questions/31836104/pyinstaller-and-onefile-how-to-include-an-image-in-the-exe-file
try:
    IconBasePath = sys._MEIPASS
except Exception:
    IconBasePath = os.path.abspath(".")
ExeIcon = Image.open(os.path.join(IconBasePath, "VRChatObsMicMuteLink.ico"))    # exeアイコン

# タスクトレイアイコンを右クリックしたときのメニュー設定
Menu = Menu(
    MenuItem("終了", Push_Exit))

# タスクトレイに常駐
Icon = Icon(name="IconName", icon=ExeIcon, title="VRChat OBS マイクミュート連携", menu=Menu)
Icon.run()
