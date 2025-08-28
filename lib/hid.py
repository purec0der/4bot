"""
HID (ヒューマンインターフェースデバイス)の操作系
主にキーボード、マウス
"""
from lib.ic import interception
from lib.stroke import *
import time
import math
import random
from ctypes import *

user32 = cdll.LoadLibrary("user32.dll")
imm32 = cdll.LoadLibrary("Imm32.dll")

# POINT構造体
class _POINT(Structure):
  _fields_ = (
    ('x', c_long),
    ('y', c_long),
  )

# Define constants
VK_NUMLOCK = 0x90
VK_CAPITAL = 0x14
VK_SCROLL = 0x91
WM_IME_CONTROL = 0x0283
IMC_GETCONVERSIONMODE = 0x0001
IME_CMODE_KATAKANA = 0x0002
IMC_GETOPENSTATUS = 0x0005

# Define toggle keys
class ToggleKey:
  TGL_NUMLOCK = 10000
  TGL_CAPSLOCK = 10001
  TGL_SCROLLLOCK = 10002
  TGL_KANALOCK = 10003
  TGL_IME = 10004

class Hid:
  # debug
  DEBUG = False

  # UWSCのマウスボタンに対応
  LEFT = 0
  RIGHT = 1
  WHEEL = 5

  # UWSCのキー入力の定数に対応
  CLICK = 0
  DOWN = 1
  UP = 2

  def __init__(self,key_input,mouse_input):
    """
    初期化
    :param key_input: interception経由のkey strokeのインスタンス
    :param mouse_input: interception経由のmouse strokeのインスタンス
    """
    self.ki = key_input
    self.mi = mouse_input

    #keyの最終CLICK時間を設定
    self.key_last_click = [float(0.0)] * 255
    #key down check
    self.key_down_flag = [0] * 255

  @staticmethod
  def load(khwid,mhwid):
    """
    必要なインスタンスをすべて生成して自身のインスタンスを返却する
    :pram khwid: <string> キーボードのHWID
    :pram mhwid: <string> マウスのHWID
    """
    # interceptionのインスタンス生成してhwid取得
    cinterception = interception()
    kdevice = cinterception.search_device_handle(khwid)
    mdevice = cinterception.search_device_handle(mhwid)
    # キーボードマウスのインスタンス生成
    mi = mouseInput(cinterception,mdevice)
    ki = keyInput(cinterception,kdevice)

    # すべてセットされた自身のインスタンスを返却
    return Hid(ki,mi)


  def kbd(self,key,cmd,wait=0,no_wait=False):
    """
    キーボードの操作
    :param key: <string> stroke.pyのSC_DECIMALのキー
    :param cmd: <int> CLICK,DOWN,UP の定数
    :param wait: <int> 実行までの待機時間 [ms]
    :param no_wait: <bool> 前回のキー入力した時間からの待機処理をしないようにする
    """
    
    if (cmd == self.CLICK):
      # キーダウンをしていない場合にのみキークリックを実行
      if (self.key_down_flag[SC_DECIMAL[key]] != self.DOWN):
        if (self.DEBUG): print(key + ' : CLICK')
        #キーの待機 (前回のクリックから早い場合待機をいれる 現状 0.1)
        if (no_wait == False):
          while (time.time() < self.key_last_click[SC_DECIMAL[key]] + 0.1 - (wait/1000)):
            time.sleep(0.01)

        self.ki.click(key,wait)

        # キークリックの最終時間保存
        self.key_last_click[SC_DECIMAL[key]] = time.time()
    elif (cmd == self.DOWN):
      # キーダウンをしていない場合にのみキーダウン実行
      if (self.key_down_flag[SC_DECIMAL[key]] != self.DOWN):
        if (self.DEBUG): print(key + ' : DOWN')
        self.ki.down(key,wait)
        self.key_down_flag[SC_DECIMAL[key]] = self.DOWN
    elif (cmd == self.UP):
      # キーダウンをしている場合にのみキーアップ実行
      if (self.key_down_flag[SC_DECIMAL[key]] != self.UP):
        if (self.DEBUG): print(key + ' : UP')
        self.ki.up(key,wait)
        self.key_down_flag[SC_DECIMAL[key]] = self.UP

  def all_key_up(self):
    """
    すべてのキーダウンを解除
    """
    for i in range(len(self.key_down_flag)):
      key_value = SC_DECIMAL.get(i)
      if (key_value is not None):
        if (self.key_down_flag[i] == self.DOWN):
          self.kbd(i, self.UP,1)

  def mmv(self,x,y,wait=0):
    """
    マウスを指定位置まで移動
    """
    self.mi.move_absolute(x,y,wait)

  def btn(self,btn,stat,x=-1,y=-1,wait=0):
    """
    マウスの操作
    :param btn: <int> LEFT,RIGHT,WHEEL の定数
    :param cmd: <int> CLICK,DOWN,UP の定数 WHEELの場合はスクロール
    :param x: <int> x座標
    :param y: <int> y座標
    :param wait: <int> 実行までの待機時間 [ms]
    """
    if (btn == self.LEFT):
      if (stat == self.CLICK):
        self.mi.left_click(x,y,wait)
      elif (stat == self.DOWN):
        self.mi.left_down(x,y,wait)
      elif (stat == self.UP):
        self.mi.left_up(x,y,wait)

    elif (btn == self.RIGHT):
      if (stat == self.CLICK):
        self.mi.right_click(x,y,wait)
      elif (stat == self.DOWN):
        self.mi.right_down(x,y,wait)
      elif (stat == self.UP):
        self.mi.right_up(x,y,wait)

    elif (btn == self.WHEEL):
      scroll = stat * 200
      if (scroll < -10000):
        scroll = -10000
      if (10000 < scroll):
        scroll = -10000
      self.mi.wheel(scroll,wait)

  def get_mouse_xy(self):
    """
    マウスの座標を返す
    :return: {x: y:} のlist
    """
    point = _POINT()
    user32.GetCursorPos(byref(point))
    return {'x':point.x, 'y':point.y}

  def input_hiragana(self,input_string):
    """
    ひらがなの文字列をキー入力する
    :param input_string: 文字列 "あいうえお" みたいな
    """
    # ひらがなを配列に変換
    input_key_array = self.encode_hiragana_to_char(input_string)
    # 入力
    for i in range(len(input_key_array)):
      self.kbd(input_key_array[i], self.CLICK, 30)

  def encode_hiragana_to_char(self,japanese_str):
    """
    ひらがな文字列を1文字ずつ分解して配列にする
    :param japanese_str: 文字列 "あいうえお" みたいな
    :return: 分割した配列 ['a','i','u','e','o'] みたいな
    """
    conversion_map = {
      'あ': 'a', 'い': 'i', 'う': 'u', 'え': 'e', 'お': 'o',
      'か': ['k', 'a'], 'き': ['k', 'i'], 'く': ['k', 'u'], 'け': ['k', 'e'], 'こ': ['k', 'o'],
      'が': ['g', 'a'], 'ぎ': ['g', 'i'], 'ぐ': ['g', 'u'], 'げ': ['g', 'e'], 'ご': ['g', 'o'],
      'さ': ['s', 'a'], 'し': ['s', 'hi'], 'す': ['s', 'u'], 'せ': ['s', 'e'], 'そ': ['s', 'o'],
      'ざ': ['z', 'a'], 'じ': ['z', 'i'], 'ず': ['z', 'u'], 'ぜ': ['z', 'e'], 'ぞ': ['z', 'o'],
      'た': ['t', 'a'], 'ち': ['t', 'i'], 'つ': ['ts', 'u'], 'て': ['t', 'e'], 'と': ['t', 'o'],
      'だ': ['d', 'a'], 'ぢ': ['d', 'i'], 'づ': ['d', 'u'], 'で': ['d', 'e'], 'ど': ['d', 'o'],
      'な': ['n', 'a'], 'に': ['n', 'i'], 'ぬ': ['n', 'u'], 'ね': ['n', 'e'], 'の': ['n', 'o'],
      'は': ['h', 'a'], 'ひ': ['h', 'i'], 'ふ': ['f', 'u'], 'へ': ['h', 'e'], 'ほ': ['h', 'o'],
      'ば': ['b', 'a'], 'び': ['b', 'i'], 'ぶ': ['b', 'u'], 'べ': ['b', 'e'], 'ぼ': ['b', 'o'],
      'ぱ': ['p', 'a'], 'ぴ': ['p', 'i'], 'ぷ': ['p', 'u'], 'ぺ': ['p', 'e'], 'ぽ': ['p', 'o'],
      'ま': ['m', 'a'], 'み': ['m', 'i'], 'む': ['m', 'u'], 'め': ['m', 'e'], 'も': ['m', 'o'],
      'や': ['y', 'a'], 'ゆ': ['y', 'u'], 'よ': ['y', 'o'],
      'ら': ['r', 'a'], 'り': ['r', 'i'], 'る': ['r', 'u'], 'れ': ['r', 'e'], 'ろ': ['r', 'o'],
      'わ': ['w', 'a'], 'を': ['w', 'o'], 'ん': ['n', 'n'],
    }
    # 変換結果を格納するリスト
    halfwidth_array = []
    for char in japanese_str:
      if char in conversion_map:
        halfwidth_array.extend(conversion_map[char])  # リストを展開して追加
      else:
        continue
    return halfwidth_array

  def get_key_state(self,code):
    """
    UWSCのGETKEYSTATEと多分同じ
    code は これだけVK
    """
    hwnd = user32.GetForegroundWindow()
    toggle = False

    if code in [ToggleKey.TGL_NUMLOCK, ToggleKey.TGL_CAPSLOCK, ToggleKey.TGL_SCROLLLOCK]:
      if code == ToggleKey.TGL_NUMLOCK:
        code = VK_NUMLOCK
      elif code == ToggleKey.TGL_CAPSLOCK:
        code = VK_CAPITAL
      elif code == ToggleKey.TGL_SCROLLLOCK:
        code = VK_SCROLL
      toggle = True
    else:
      hime = imm32.ImmGetDefaultIMEWnd(hwnd)
      if code == ToggleKey.TGL_KANALOCK:
        mode = user32.SendMessageW(hime, WM_IME_CONTROL, IMC_GETCONVERSIONMODE, 0)
        return (mode & IME_CMODE_KATAKANA) > 0
      elif code == ToggleKey.TGL_IME:
        state = user32.SendMessageW(hime, WM_IME_CONTROL, IMC_GETOPENSTATUS, 0)
        return state != 0

    key_state = user32.GetKeyState(code)
    if toggle:
      return (key_state & 0x0001) > 0
    else:
      return (key_state & 0x8000) > 0


  ####################
  # それっぽいマウス移動
  ####################

  def rmmv(self,x,y,wait=0):
    """
    移動の軌跡を作ってマウス移動
    """
    time.sleep(wait/1000)

    x = x + int(random.uniform(0,5))
    y = y + int(random.uniform(0,5))
    self.mouse_move_real(x,y,int(random.uniform(1,4)))

    res = [x,y]
    return res

  def rbtn(self,key,type,x,y,wait):
    """
    移動の軌跡を作ってマウスアクション
    """
    time.sleep(wait/1000)
    x = x + random.uniform(0,5)
    y = y + random.uniform(0,5)
    self.mouse_move_real(x,y,int(random.uniform(1,4)))

    self.btn(key,type,x,y)

  def mouse_move_real(self, x, y, mspeed):
    mouse_xy = self.get_mouse_xy()
    while False == (abs(mouse_xy['x'] - x) <= 3 and abs(mouse_xy['y'] - y) <= 3):
      x_len = abs(mouse_xy['x'] - x)
      y_len = abs(mouse_xy['y'] - y)
      move_len = math.sqrt( pow(x_len,2) + pow(y_len,2))
      x_speed = self.get_mouse_speed(move_len + 1,mspeed)
      y_speed = self.get_mouse_speed(move_len + 1,mspeed)

      x_move = random.uniform(0,x_speed) + 1
      y_move = random.uniform(0,y_speed) + 1
      if (x_len != 0):
        y_move = int(y_move * (y_len / x_len))
      else:
        y_move = int(y_move * (y_len / 1))


      if (x_len < x_move):
        x_move = x_len

      if (y_len < y_move):
        y_move = y_len

      if (x == mouse_xy['x'] and y_move == 0):
        y_move = 1

      if (x < mouse_xy['x']):
        x_move = 0 - x_move

      if (y < mouse_xy['y']):
        y_move = 0 - y_move

      self.mmv(mouse_xy['x'] + x_move, mouse_xy['y'] + y_move,1)

      mouse_xy = self.get_mouse_xy()

  def get_mouse_speed(self,move_len,mspeed):
    return int(move_len / mspeed) + 1