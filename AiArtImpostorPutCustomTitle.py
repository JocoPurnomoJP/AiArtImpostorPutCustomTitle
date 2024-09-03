#py AiArtImpostorPutCustomTitle.py
#py -m PyInstaller AiArtImpostorPutCustomTitle.py --clean --onefile --icon=AiArtImpostorPutCustomTitle.ico --add-data AiArtImpostorPutCustomTitle.ico:. --noconsole
#↓コンソールが不要な場合は以下を選択。コンソールを残したままにしたい場合↑でEXEを作る
#py -m PyInstaller AiArtImpostorPutCustomTitle.py --clean --onefile --noconsole
from pywinauto import Desktop
from pywinauto.application import Application
from pywinauto import timings
timings.Timings.window_find_timeout = 1

import tkinter as tk
from tkinter import ttk #frameのため必要
from tkinter.font import Font
import tkinter.filedialog
import tkinter.messagebox as messagebox

#from pywinauto import timings
#from pywinauto import keyboard #use send_keys
import os
import time
import sys
#import traceback
import csv

import ctypes #AIArtImpostorの位置を取得
from ctypes.wintypes import HWND, DWORD, RECT

import pyautogui #pywiunautoでオブジェクトが取れないので、座標操作をする
#https://qiita.com/run1000dori/items/301bb63c8a69c3fcb1bd

import pyperclip #コピペ用
import math
from bisect import bisect_left #近似値を求める時につかう

#https://github.com/studio-ousia/mojimoji
#py -m pip install mojimoji
import mojimoji #半角文字＞全角文字に変換

#debug
#import pprint

#ログなどの差別化のため日付時刻取得
#https://note.nkmk.me/python-datetime-usage/
import datetime
dt = datetime.datetime.now()
dateStr = dt.strftime('%Y%m%d%H%M%S')

#フルスクリーン取得用
from ctypes import windll
import win32gui
user32 = windll.user32
user32.SetProcessDPIAware() # optional, makes functions return real pixel numbers instead of scaled values
full_screen_rect = (0, 0, user32.GetSystemMetrics(0), user32.GetSystemMetrics(1))

#定数 最下行右下はわからない固定
UNKNOWN_TITLE = "わからない"
#１８項目で終わりなので、行数は９行で終わり
MAX_ROW = 9
# ゲーム側の入力項目は６行＊３列
MAX_ROWGAME = 6

# わからないが入る位置は１８番目
LAST_TITLE_INDEX = 18

#ゲームタイトル
GAME_TITLE = "AIArtImpostor"
#タイトル名とは別にスクリーン名があるので、その定数
SCREEN_TITLE_EXE = GAME_TITLE + ".exe"
SCREEN_TITLE_NOEXE = GAME_TITLE

#待ち時間共通
MIN_WAIT = 0.1

#最大文字数、文字幅制御あるかわからないが、全角で16文字限界、半角も同じ。AIに生成させた場合はそれ以上入る
#なおカテゴリーも同じだった
MAX_WORDS = 16

#画面座標系定数、こちらは画面左上からの相対距離
#2024/07/08 ゲーム内の解像度を参照して、オブジェクトの間隔などを可変させるので定数から変数にする
#おそらくだがゲーム自体がディスプレイの解像度によって、サイズが異なる可能性大
#そうなると、補正値を入力しないと動かないかも
#カテゴリー入力項目 TAPLE形式から配列へ
#CATEGORY_TXT_POS = [600,175]
#TITLE_TXT_POS = (275,175)
CATEGORY_TXT_X = 630
CATEGORY_TXT_Y = 175
wCATEGORY_TXT_X = CATEGORY_TXT_X
wCATEGORY_TXT_Y = CATEGORY_TXT_Y

#https://crystage.com/contents/resolution/
#以下は 1280 * 720の場合は動作する大きさ　　HDTV 	1280×720 	16:9
TITLE_WIDTH_DISTANCE = 320
TITLE_HEIGHT_DISTANCE = 67
wTITLE_WIDTH_DISTANCE = TITLE_WIDTH_DISTANCE
wTITLE_HEIGHT_DISTANCE = TITLE_HEIGHT_DISTANCE

#以下は 1600 * 900の場合は動作する大きさ　　WXGA++ 	Wide XGA++ 	1600×900　16:9
#1.25倍で増える
#TITLE_WIDTH_DISTANCE = 375
#TITLE_HEIGHT_DISTANCE = 87

#以下は 1920 * 1080の場合は動作する大きさ　　FHD  	1920×1080　16:9
#TITLE_WIDTH_DISTANCE = 450
#TITLE_HEIGHT_DISTANCE = 105

#QWXGA 	Quad-Wide-XGA 	2048×1152 	2,359,296 	16:9 
#TITLE_WIDTH_DISTANCE = 480
#TITLE_HEIGHT_DISTANCE = 112

#WQHD (Wide-Quad-HD) 	2560×1440 	16:9
#TITLE_WIDTH_DISTANCE = 600
#TITLE_HEIGHT_DISTANCE = 140

#基準値を基に倍率をセットする
RATE_LIST = [
1,
1.33,
1.575,
#1.6,
2
]
#上と対になる画面サイズ
WINDOW_LIST = [
2000, #1280 * 720
2500, #1600 * 900
3000, #1920 * 1080
#3200, #2048 * 1152
4000 #2560 * 1440
]

#共通変数
#ボタンを押すたびにコロコロ画面が動いていたのでボタン押下直後の画面サイズを記録し
#処理終了後にサイズをセットして、画面再構成することで、画面がちらついたり、動かなくなった
#root.resizableも試したが、画面がちらついたので却下
wWidth = 0
wHeight = 0

#環境によって、Taskbarに記載されるゲーム名が「AIArtImpostor.exeだったり、AIArtImposterと.exeがなかったりするのでその対策」
currentScreenName = SCREEN_TITLE_EXE

#お題を保持している変数
wCategory = ""
wTitles = [] #タイトルを単純に配列化

def resource_path(relative_path):
    try:
        #Retrieve Temp Path
        base_path = sys._MEIPASS
    except Exception:
        #Retrieve Current Path Then Error 
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def keepWindowSize():
    #画面のサイズ変更を防ぐメソッド、実行時にサイズが０なら現在のサイズを控える
    #そうでなければ、控えた値を使って、画面サイズを調整し、控え値を０に戻す
    global wWidth
    global wHeight
    if wWidth == 0:
        wWidth = root.winfo_width()
        wHeight = root.winfo_height()
    else:
        root.geometry(f"{wWidth}x{wHeight}")
        wWidth = 0
        wHeight = 0
    #高さも横幅に連動しているので、横幅だけチェックで良い

#先頭と末尾のダブルクォーてションやクォテーション削除
#ついでに半角スペースなども再度チェックする
#既存の関数を使うと真ん中のダブルクォートなども外れてしまうので
def stripQuote(str):
    str = str.strip()
    if len(str) <= 0:
        return str
    if str[0:1] == "'" or str[0:1] == '"' or str[0:1] == "’" or str[0:1] == "”":
        str = str[1:]
    if str[-1] == "'" or str[-1] == '"' or str[-1] == "’" or str[-1] == "”":
        str = str[0:-1]
    return str.strip()

#文字数を超えている場合は該当項目を赤文字にして、そうではない場合は黒文字へ戻す
def changeOverWordsTextColor(event=None):
    #チェックしていく、文字色を変える
    for i in range(MAX_ROW):
        #タグは再生性するので毎回削除
        for tag in txts_col1[i].tag_names():
            txts_col1[i].tag_delete(tag)
            
        if len(stripQuote(txts_col1[i].get(1.0, tk.END))) > MAX_WORDS:
            txts_col1[i].tag_config("color2", foreground="red")
            txts_col1[i].tag_add("color2",1.0, tk.END)
        else:
            txts_col1[i].tag_config("color1", foreground="black")
            txts_col1[i].tag_add("color1",1.0, tk.END)
    for i in range(MAX_ROW):
        #タグは再生性するので毎回削除
        for tag in txts_col2[i].tag_names():
            txts_col2[i].tag_delete(tag)
            
        if len(stripQuote(txts_col2[i].get(1.0, tk.END))) > MAX_WORDS:
            txts_col2[i].tag_config("color2", foreground="red")
            txts_col2[i].tag_add("color2",1.0, tk.END)
        else:
            txts_col2[i].tag_config("color1", foreground="black")
            txts_col2[i].tag_add("color1",1.0, tk.END)
    #カテゴリーもチェック
    #タグは再生性するので毎回削除
    for tag in txt_category.tag_names():
        txt_category.tag_delete(tag)
    if len(stripQuote(txt_category.get(1.0, tk.END))) > MAX_WORDS:
        txt_category.tag_config("color2", foreground="red")
        txt_category.tag_add("color2",1.0, tk.END)
    else:
        txt_category.tag_config("color1", foreground="black")
        txt_category.tag_add("color1",1.0, tk.END)
        
#座標取得
#https://qiita.com/ShortArrow/items/409f9695c458433d0744
def GetWindowRectFromName(TargetWindowTitle:str)-> tuple:
    TargetWindowHandle = ctypes.windll.user32.FindWindowW(0, TargetWindowTitle)
    Rectangle = ctypes.wintypes.RECT()
    ctypes.windll.user32.GetWindowRect(TargetWindowHandle, ctypes.pointer(Rectangle))
    return (Rectangle.left, Rectangle.top, Rectangle.right, Rectangle.bottom)

#リスト内から最も近い値を取得して、要素位置を返す
#https://qlitre-weblog.com/how-to-get-nearest-value-in-list-python
def get_nearest_value_in_list(an_iterable: list, target: int) -> int:
    # 最初にソートする
    an_iterable.sort()
    # ターゲットから挿入されるべきインデックスを取得
    index = bisect_left(an_iterable, target)
    #debug
    #print(f"insert:{index}")

    # indexが先頭の場合は先頭の要素
    if index == 0:
        return index
    #    return an_iterable[0]
    # indexが末尾の場合は末尾の要素
    if index == len(an_iterable):
        return index - 1
    #    return an_iterable[-1]
    # 前後を比較してtargetに近い方を返す
    a = math.fabs(target - an_iterable[index - 1])
    b = math.fabs(target - an_iterable[index])
    if a <= b:
        return index - 1
        #return an_iterable[index - 1]
    else:
        return index
        #return an_iterable[index]

def on_clear():
    global wTitles
    wTitles = []
    keepWindowSize()
    #print("on_clear")
    #print(root.winfo_width())
    #for loopで削除する
    for i in range(MAX_ROW):
        txt1 = txts_col1[i]
        txt2 = txts_col2[i]
        txt1.delete(1.0, tk.END)
        txt2.delete(1.0, tk.END)
        
    txt_category.delete(1.0, tk.END)
    putUnknown()
    keepWindowSize()

#右下最下行は必ず固定値
def putUnknown():
    txts_col2[MAX_ROW - 1].delete(1.0, tk.END)
    txts_col2[MAX_ROW - 1].insert(0., UNKNOWN_TITLE)

#実際の処理メソッド
#csvからテキストへお題を入力する
def import_from_csv():
    global wTitles
    #print("import_from_csv")
    
    #配列クリア
    wTitles = []
    
    #ファイルダイアログから
    fTyp = [("csvファイル", "*.csv")]
    iDir = os.getcwd() 
    # 上のファイルパスでハマった。_MEIのシステムパスが __file__のときは勝手に設定されてしまう
    #参考　https://qiita.com/Authns/items/f3cf6e9f27fc0b5632b3
    file_name = tkinter.filedialog.askopenfilename(filetypes=fTyp, initialdir=iDir)
    #ファイル名空欄の場合は処理終了。いわゆるキャンセルした時の処理
    if len(file_name) == 0:
        return
    
    #入力文字列全クリア
    on_clear()
    wTitles = []
    
    basename = os.path.basename(file_name)
    
    #csv読み込み開始 エラー処理入れる
    try:
        with open(file_name ,encoding=combobox_encode.get()) as f:
            reader = csv.reader(f)
            for row in reader:
                wTitles.append(stripQuote(row[0].strip()))
        label_memo2.config(text="CSVを読込、内容を各項目に反映しました。", foreground="black")
    except:
        label_memo2.config(text="CSVの読込に失敗しました。ファイルの権限や文字コードを見直して下さい。", foreground="red")
    
    #txtに入れていく 左列から埋めて、その後右列
    adjustTitles(True)
    
    #カテゴリーセット
    txt_category.insert(0., basename[0:basename.rfind(".")])
    
    #文字数チェック
    changeOverWordsTextColor()
    
#ゲーム画面の大きさを把握して補正値を返す
def getBaseDistance(sizeNum):
    global wCATEGORY_TXT_X
    global wCATEGORY_TXT_Y
    global wTITLE_WIDTH_DISTANCE
    global wTITLE_HEIGHT_DISTANCE
    
    index = get_nearest_value_in_list(WINDOW_LIST, sizeNum)
    
    #debug
    #print(f"screenSize:{sizeNum}")
    #print(f"Rate:{RATE_LIST[index]}")
    
    #倍率に応じて補正する
    wCATEGORY_TXT_X = math.floor(CATEGORY_TXT_X * RATE_LIST[index])
    wCATEGORY_TXT_Y = math.floor(CATEGORY_TXT_Y * RATE_LIST[index])
    wTITLE_WIDTH_DISTANCE = math.floor(TITLE_WIDTH_DISTANCE * RATE_LIST[index])
    wTITLE_HEIGHT_DISTANCE = math.floor(TITLE_HEIGHT_DISTANCE * RATE_LIST[index])
    #お題の高さは更に補正する
    if index == 1:
        wTITLE_HEIGHT_DISTANCE = wTITLE_HEIGHT_DISTANCE - 8
    elif index == 2:
        wTITLE_HEIGHT_DISTANCE = wTITLE_HEIGHT_DISTANCE - 4
        
#ゲーム画面を取得してトップフォーカスする
#成功時はTrue、失敗時はFalseを返す
def setFocusGameWindow():
    global currentScreenName
    #AiArtImposterから画面取得
    #print("1:" + str(datetime.datetime.now()))
    app = Application(backend='uia')
    #print("2:" + str(datetime.datetime.now()))
    desktop = Desktop(backend='uia')
    #print("3:" + str(datetime.datetime.now()))
    #app = Application(backend='win32')
    try:
        #デスクトップのツールバーでちっちゃくなっている場合はappだけだと取得できないので
        #ツールバーをクリックする処置をいれる
        taskbar = desktop.window(class_name="Shell_TrayWnd")
        #print("4:" + str(datetime.datetime.now()))
        #debug
        #taskbar.print_control_identifiers()
        window_button = taskbar.child_window(title_re=currentScreenName + " -*", control_type="Button")
        window_button.click_input()
        #print("5:" + str(datetime.datetime.now()))
        app.connect(title=GAME_TITLE)
        #print("6:" + str(datetime.datetime.now()))
    except:
        #１度失敗した場合は別の定数でも探してみる
        #なおスクリーン名に３つ目が出たら、別メソッド、配列にして対応予定
        try:
            if currentScreenName == SCREEN_TITLE_EXE:
                currentScreenName = SCREEN_TITLE_NOEXE
            else:
                currentScreenName = SCREEN_TITLE_EXE
            taskbar = desktop.window(class_name="Shell_TrayWnd")
            window_button = taskbar.child_window(title_re=currentScreenName + " -*", control_type="Button")
            window_button.click_input()
            app.connect(title=GAME_TITLE)
        except:
            messagebox.showinfo('画面取得失敗', 'ゲーム画面が見つかりません。ゲームを起動して、カスタムお題の画面を開いてから実行してください') 
            return False
        
    #フォーカスをゲーム画面に当てる
    #dlg = app.top_window()
    top_window = app.top_window()
    #print("7:" + str(datetime.datetime.now()))
    #フルスクリーンの時はset.focusできない
    #フルスクリーンか確認
    tpl = GetWindowRectFromName(GAME_TITLE)
    #print("8:" + str(datetime.datetime.now()))
    #debug
    #print(f"full_screen_rect[2]:{full_screen_rect[2]}")
    #print(f"tpl[2]:{tpl[2]}")
    #print(f"full_screen_rect[3]:{full_screen_rect[3]}")
    #print(f"tpl[3]:{tpl[3]}")
    #if full_screen_rect[2] == tpl[2] and full_screen_rect[3] == tpl[3]:
    if full_screen_rect[2] != tpl[2] or full_screen_rect[3] != tpl[3]:
        top_window.restore().set_focus()
        #debug
        #messagebox.showinfo('画面取得失敗', 'ゲーム画面がフルスクリーンでは本ツールは機能しません。\nオプションからウィンドウモードの画面サイズに変更してください。') 
        #return False
    #print("9:" + str(datetime.datetime.now()))
    #以下でフォーカスはセットできたので、画面自体は正常に取得できている
    #descendants = dlg.descendants()
    #pprint.pprint(descendants)
    #best_matches = findbestmatch.find_best_control_matches("TitleBar", descendants)
    #best_matches[0].set_focus()
    time.sleep(MIN_WAIT)
    
    return True

#メイン機能、テキストに写っているものをゲーム画面に反映
def copy_to_screen():
    #print("copy_to_screen")
    #print("10:" + str(datetime.datetime.now()))
    if setFocusGameWindow() == False:
        return
    #print("11:" + str(datetime.datetime.now()))
    
    try:
        #debug
        #print(GetWindowRectFromName(GAME_TITLE)) #右上に強引に持ってたら　私の環境だと(1024, 0, 2320, 759)になった
        gameWindowPosTpl = GetWindowRectFromName(GAME_TITLE)
        #共通変数の値を調整する
        getBaseDistance(gameWindowPosTpl[2] - gameWindowPosTpl[0] + gameWindowPosTpl[3] - gameWindowPosTpl[1])
        #print("12:" + str(datetime.datetime.now()))
        #debug　試しにカテゴリーに値が入るかやってみる
        #print(f"X:{wCATEGORY_TXT_X}")
        #print(f"Y:{wCATEGORY_TXT_Y}")
        
        #マウスをカテゴリーのテキスト入力まで移動
        #print(f"x:{gameWindowPosTpl[0] + CATEGORY_TXT_POS[0]}, y:{gameWindowPosTpl[1] + CATEGORY_TXT_POS[1]}")
        categoryX,categoryY = gameWindowPosTpl[0] + wCATEGORY_TXT_X,gameWindowPosTpl[1] + wCATEGORY_TXT_Y
        pyautogui.moveTo(categoryX,categoryY)
        pyautogui.click(categoryX,categoryY)
        #print("13:" + str(datetime.datetime.now()))
        #https://qiita.com/umashikate/items/98c94cdd269ea26c41c6
        #CTRL + A で、まずカテゴリーを消す
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(MIN_WAIT)
        pyautogui.press('delete')
        time.sleep(MIN_WAIT)
        
        #1文字ごとに0.25秒の間隔で入力
        #pyautogui.write('Hello world!', interval=0.25)
        #これでコピペ出来る
        pyperclip.copy(txt_category.get(1.0, tk.END).strip()) #こっちがクリップボード
        pyautogui.hotkey('ctrl', 'v') #こっちがペーストのホットキー
        
        #ということで、後は６行＊３列に対して、ぐるぐる回すことにする
        adjustTitles(False)
        
        #カテゴリーの位置から計算して、左上から入力していく
        for i in range(MAX_ROWGAME):
             pyperclip.copy(wTitles[i * 3])
             pyautogui.moveTo(categoryX - wTITLE_WIDTH_DISTANCE, categoryY + wTITLE_HEIGHT_DISTANCE * (i + 1))
             pyautogui.click(categoryX - wTITLE_WIDTH_DISTANCE, categoryY + wTITLE_HEIGHT_DISTANCE * (i + 1))
             pyautogui.hotkey('ctrl', 'v')
             
             pyperclip.copy(wTitles[i * 3 + 1])
             pyautogui.moveTo(categoryX, categoryY + wTITLE_HEIGHT_DISTANCE * (i + 1))
             pyautogui.click(categoryX, categoryY + wTITLE_HEIGHT_DISTANCE * (i + 1))
             pyautogui.hotkey('ctrl', 'v')
             
             #18番目は入力できないので飛ばす
             if i * 3 + 2 < LAST_TITLE_INDEX - 1:
                 pyperclip.copy(wTitles[i * 3 + 2])
                 pyautogui.moveTo(categoryX + TITLE_WIDTH_DISTANCE, categoryY + wTITLE_HEIGHT_DISTANCE * (i + 1))
                 pyautogui.click(categoryX + TITLE_WIDTH_DISTANCE, categoryY + wTITLE_HEIGHT_DISTANCE * (i + 1))
                 pyautogui.hotkey('ctrl', 'v')
        
        #項目が編集中のままだと、全体に反映されないので、故意にカテゴリーをクリックしておく
        pyautogui.moveTo(categoryX,categoryY)
        pyautogui.click(categoryX,categoryY)
        label_memo2.config(text="各項目からゲーム画面へコピーしました。", foreground="black")
    except:
        label_memo2.config(text="ゲーム画面操作でエラーが発生しました。お題などが読み込める状態で再実行して下さい。", foreground="red")
    
#ゲーム画面からテキストへ反映する
def import_from_screen():
    global wTitles
    #print("import_from_screen")
    on_clear()
    
    if setFocusGameWindow() == False:
        return
    
    #エラー発生時は失敗扱いとする
    try:
        #debug
        #print(GetWindowRectFromName(GAME_TITLE)) #右上に強引に持ってたら　私の環境だと(1024, 0, 2320, 759)になった
        gameWindowPosTpl = GetWindowRectFromName(GAME_TITLE)
        #共通変数の値を調整する
        getBaseDistance(gameWindowPosTpl[2] - gameWindowPosTpl[0] + gameWindowPosTpl[3] - gameWindowPosTpl[1])
        
        #debug
        #マウスをカテゴリーのテキスト入力まで移動
        #print(f"x:{gameWindowPosTpl[0] + CATEGORY_TXT_POS[0]}, y:{gameWindowPosTpl[1] + CATEGORY_TXT_POS[1]}")
        categoryX,categoryY = gameWindowPosTpl[0] + wCATEGORY_TXT_X,gameWindowPosTpl[1] + wCATEGORY_TXT_Y
        pyautogui.moveTo(categoryX,categoryY)
        pyautogui.click(categoryX,categoryY)
        
        #https://qiita.com/umashikate/items/98c94cdd269ea26c41c6
        #CTRL + A で、まずカテゴリーを消す
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(MIN_WAIT)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(MIN_WAIT)
        
        #カテゴリーテキストへ反映
        txt_category.insert(0.,stripQuote(pyperclip.paste()))
        
        #カテゴリーの位置から計算して、左上からコピーしていく
        wTitles = []
        for i in range(MAX_ROWGAME):
             pyautogui.moveTo(categoryX - wTITLE_WIDTH_DISTANCE, categoryY + wTITLE_HEIGHT_DISTANCE * (i + 1))
             pyautogui.click(categoryX - wTITLE_WIDTH_DISTANCE, categoryY + wTITLE_HEIGHT_DISTANCE * (i + 1))
             pyautogui.hotkey('ctrl', 'c')
             wTitles.append(stripQuote(pyperclip.paste()))
             
             pyautogui.moveTo(categoryX, categoryY + wTITLE_HEIGHT_DISTANCE * (i + 1))
             pyautogui.click(categoryX, categoryY + wTITLE_HEIGHT_DISTANCE * (i + 1))
             pyautogui.hotkey('ctrl', 'c')
             wTitles.append(stripQuote(pyperclip.paste()))
             
             #18番目は入力できないので飛ばす
             if i * 3 + 2 < LAST_TITLE_INDEX - 1:
                 pyautogui.moveTo(categoryX + wTITLE_WIDTH_DISTANCE, categoryY + wTITLE_HEIGHT_DISTANCE * (i + 1))
                 pyautogui.click(categoryX + wTITLE_WIDTH_DISTANCE, categoryY + wTITLE_HEIGHT_DISTANCE * (i + 1))
                 pyautogui.hotkey('ctrl', 'c')
                 wTitles.append(stripQuote(pyperclip.paste()))
        
        #LAST_TITLE_INDEX分入っていない場合は空文字を入れる
        if len(wTitles) < LAST_TITLE_INDEX:
            for i in range(LAST_TITLE_INDEX - len(wTitles)):
                wTitles.append("")
        #配列からテキストへ反映
        adjustTitles(True)
        label_memo2.config(text="ゲーム画面から各項目へコピーしました。", foreground="black")
    except:
        label_memo2.config(text="ゲーム画面操作でエラーが発生しました。お題などが読み込める状態で再実行して下さい。", foreground="red")
    
    #テキスト内の文字数チェック
    changeOverWordsTextColor()

#現在の画面からCSVへ
def export_to_csv():
    global wTitles
    #print("export_to_csv")
    
    #カテゴリーをファイル名にするので、Windowsでファイル名に使えない文字は全角文字に置換
    #￥（円マーク）
    #/（斜線、スラッシュ）
    #: （コロン)
    #*（アスタリスク）
    #?（ 疑問符、クエスチョンマーク）
    #“（ダブルコーテーション）
    #<>（不等号）
    #|（縦線、パイプライン）
    replaceChars = ('\\','/',':','*','?','"','<','>','|')
    
    #カテゴリーはデフォルトファイル名とする
    fileBaseName = txt_category.get(1.0, tk.END).strip()
    
    #一致したら全角文字に置換してしまう
    for char in replaceChars:
        fileBaseName = fileBaseName.replace(char, mojimoji.han_to_zen(char))
    
    #配列クリア
    wTitles = []
    
    #txtから配列へ
    adjustTitles(False)
    
    #debug
    #print(f"カテゴリ：{fileBaseName}")
    
    #ファイル保存ダイアログ
    file_name = tkinter.filedialog.asksaveasfilename(
    title = "名前をつけて保存。カテゴリー名を使うことをお勧め",
    filetypes = [("csvファイル", "*.csv")], 
    initialfile = fileBaseName,
    initialdir = os.getcwd(),
    defaultextension = "csv"
    )
    #ファイル名空欄の場合は処理終了。いわゆるキャンセルした時の処理
    if len(file_name) == 0:
        return
        
    #csv書き出し
    #なお、９行目までしか読み込まない。１行目からデータ行として扱う
    #newline=''がないとWindows環境はバグる
    try:
        with open(file_name, 'w', encoding=combobox_encode.get(), newline='') as f:
            #writer = csv.writer(f, lineterminator='\n')
            writer = csv.writer(f)
            for row in wTitles:
                writer.writerow([row])
        #writer.writerows(wTitles)これは使えなかった。文字列分解されてしまった
        label_memo2.config(text="各項目の内容をCSVへ出力しました。", foreground="black")
    except:
        label_memo2.config(text="CSVへの出力に失敗しました。フォルダへの権限を見直して下さい。", foreground="red")

#入力されたタイトルをスクリーンに入力するように整形する
#カスタムお題作成時は３列で６行、３列目の６行目がわからない相当なので、
#２列９行のインポスター画面と合わない。なのでそれを調整する
#fromtoFlg, Trueだとテキストから配列へ、Falseだと配列からテキストへ逆反映
def adjustTitles(fromtoFlg):
    global wTitles
    #print("adjustTitles")
    
    if fromtoFlg:
        wTitles.append("")
        for i in range(MAX_ROW):
            if i < len(wTitles):
                txts_col1[i].insert(0., wTitles[i])
        for i in range(MAX_ROW):
            if i + MAX_ROW < len(wTitles):
                txts_col2[i].insert(0., wTitles[i + MAX_ROW])
        putUnknown()
    
    else:
        wTitles = []
        #左列をまず埋めてから、右列を埋めるという順番になる
        for i in range(MAX_ROW):
            wTitles.append(stripQuote(txts_col1[i].get(1.0, tk.END)))
        for i in range(MAX_ROW):
            wTitles.append(stripQuote(txts_col2[i].get(1.0, tk.END)))
    #debug
    #pprint.pprint(wTitles)
            
# Create the main window
root = tk.Tk()
root.title("Ai Art Impostor Put Custom Title ver 1.10")

# iconとEXEマークの画像
logo=resource_path('AiArtImpostorPutCustomTitle.ico') #ソースコードと画像は同じディレクトリにある前提
root.iconbitmap(default=logo)

root.minsize(100, 100)
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

# Frame
frame1 = ttk.Frame(root, padding=10)
frame1.rowconfigure(1, weight=1)
frame1.columnconfigure(0, weight=1)
frame1.grid(sticky=(tk.N, tk.W, tk.S, tk.E))

# Frameを作成すると画面そのものに配置ができなくなるみたい
# Frameの上にボタンなどを配置する形
# Frame
frame1 = ttk.Frame(root, padding=10)
frame1.rowconfigure(1, weight=1)
frame1.columnconfigure(0, weight=1)
frame1.grid(sticky=(tk.N, tk.W, tk.S, tk.E))

# Button
#ファイルダイアログからCSVを読み込ませて、各エリアに反映
button_csv = ttk.Button(
    frame1, text='CSVからこの画面へ', width=45,
    command=import_from_csv)
button_csv.config( width =45 )
button_csv.grid(
    row=0, column=0, columnspan=1, sticky=(tk.N, tk.W, tk.S, tk.E))

#ファイルダイアログからCSVを読み込ませて、各エリアに反映
button_ext = ttk.Button(
    frame1, text='この画面からゲームへコピー', width=45,
    command=copy_to_screen)
button_ext.config( width =45 )
button_ext.grid(
    row=0, column=1, columnspan=1, sticky=(tk.N, tk.W, tk.S, tk.E))

#お題のカテゴリー
f1 = Font(family='Helvetica', size=16)
txt_category = tk.Text(frame1, height=1, width=45)
txt_category.configure(font=f1)
txt_category.tag_config("color1", foreground="black")
txt_category.tag_config("color2", foreground="red")
txt_category.grid(row=1, column=0, columnspan=2, sticky=(tk.N, tk.W, tk.S, tk.E))
txt_category.insert(0., "カテゴリーの入力")
txt_category.bind('<KeyRelease>', changeOverWordsTextColor)

# Text カスタムタイトルは17箇所に入力
# 見た目をインポスターがお題を当てる時に合わせる目的で２列の９行で行う、
#２列目の最期が”わからない”固定
#ということで固定ForLoopで回す
txts_col1 = [tk.Text(frame1, height=1, width=32, undo=True) for i in range(MAX_ROW)]
txts_col2 = [tk.Text(frame1, height=1, width=32, undo=True) for i in range(MAX_ROW)]

for i in range(MAX_ROW):
    txt1 = txts_col1[i]
    txt1.configure(font=f1)
    txt1.insert(0., str(i+1) + "お題の入力")
    txt1.grid(row=i+2, column=0, columnspan=1, sticky=(tk.N, tk.W, tk.S, tk.E))
    txt1.bind('<KeyRelease>', changeOverWordsTextColor)
for i in range(MAX_ROW):
    txt2 = txts_col2[i]
    txt2.configure(font=f1)
    txt2.insert(0., str(i+MAX_ROW+1) + "お題の入力")
    txt2.grid(row=i+2, column=1, columnspan=1, sticky=(tk.N, tk.W, tk.S, tk.E))
    txt2.bind('<KeyRelease>', changeOverWordsTextColor)

#最期だけわからない固定 書き込み禁止にする
putUnknown()
txts_col2[MAX_ROW - 1].config(state=tk.DISABLED,background="gray")

# Button
#現状の画面からテキストへコピー
button_screenTotext = ttk.Button(
    frame1, text='ゲームからこの画面へコピー', width=45,
    command=import_from_screen)
button_screenTotext.config( width =45 )
button_screenTotext.grid(
    row=MAX_ROW+2, column=0, columnspan=1, sticky=(tk.N, tk.W, tk.S, tk.E))

#現状のテキスト入力をCSVにして一括出力
button_textToCSV = ttk.Button(
    frame1, text='この画面の入力をCSVへ', width=45,
    command=export_to_csv)
button_textToCSV.config( width =45 )
button_textToCSV.grid(
    row=MAX_ROW+2, column=1, columnspan=1, sticky=(tk.N, tk.W, tk.S, tk.E))

#テキストを一括クリア
button_clear = ttk.Button(
    frame1, text='Clear', width=45,
    command=on_clear)
button_clear.config( width =45 )
#なお、誤作動を防ぐため、画面下側にクリアボタンを配置する
button_clear.grid(
    row=MAX_ROW+3, column=0, columnspan=1, sticky=(tk.N, tk.W, tk.S, tk.E))

# ComboBox
#ドロップダウンリスト、CSVを読み込み・出力する時の文字コードを指定
#一応想定は0-5で
encode_fixed = ("utf_8","shift_jis","euc_jp")
combobox_encode = ttk.Combobox(frame1, state="readonly", values=encode_fixed)
combobox_encode.grid(
    row=MAX_ROW+3, column=1, columnspan=1, sticky=(tk.N, tk.W, tk.S, tk.E))
combobox_encode.set("utf_8")
#combobox_encode.get()

f2 = Font(family='Helvetica', size=9)
label_memo = ttk.Label(
    frame1, text='赤文字の項目はゲーム画面へコピーすると見切れます。CSVの読込に失敗する場合はドロップボックスのutf_8を変更して下さい。', font=f2)
label_memo.grid(
    row=MAX_ROW+4, column=0, columnspan=2, sticky=(tk.N, tk.W, tk.S, tk.E))

f3 = Font(family='Helvetica', size=12)
label_memo2 = ttk.Label(
    frame1, text='ここにシステムメッセージが表示されます', font=f3)
label_memo2.grid(
    row=MAX_ROW+5, column=0, columnspan=2, sticky=(tk.N, tk.W, tk.S, tk.E))

# Run the main event loop
root.geometry("790x425") #x * y
root.mainloop()
