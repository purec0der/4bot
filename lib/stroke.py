"""interception"""

from ctypes import *
from random import random
import time

user32 = cdll.LoadLibrary("user32.dll")


#####################################
# ストローク構造体 (Mouse / KeyBoard)
#####################################

# mouse 構造体
class _MOUSE_INPUT_DATA(Structure):
  _fields_ = (
    ('UnitId', c_ushort),
    ('Flags', c_ushort),
    ('ButtonFlags', c_ushort),
    ('ButtonData', c_ushort),
    ('RawButtons', c_ulong),
    ('LastX', c_long),
    ('LastY', c_long),
    ('ExtraInformation', c_ulong),
  )

# key 構造体
class _KEYBOARD_INPUT_DATA(Structure):
  _fields_ = (
    ('UnitId', c_ushort),
    ('MakeCode', c_ushort),
    ('Flags', c_ushort),
    ('Reserved', c_ushort),
    ('ExtraInformation', c_ulong),
  )


##########################
# マウス操作
##########################
class mouseInput:
  # InterceptionMouseState
  INTERCEPTION_MOUSE_LEFT_BUTTON_DOWN = 0x1
  INTERCEPTION_MOUSE_LEFT_BUTTON_UP = 0x2
  INTERCEPTION_MOUSE_RIGHT_BUTTON_DOWN = 0x4
  INTERCEPTION_MOUSE_RIGHT_BUTTON_UP = 0x8
  INTERCEPTION_MOUSE_MIDDLE_BUTTON_DOWN = 0x10
  INTERCEPTION_MOUSE_MIDDLE_BUTTON_UP = 0x20

  INTERCEPTION_MOUSE_BUTTON_1_DOWN = 0x1
  INTERCEPTION_MOUSE_BUTTON_1_UP = 0x2
  INTERCEPTION_MOUSE_BUTTON_2_DOWN = 0x4
  INTERCEPTION_MOUSE_BUTTON_2_UP = 0x8
  INTERCEPTION_MOUSE_BUTTON_3_DOWN = 0x10
  INTERCEPTION_MOUSE_BUTTON_3_UP = 0x20

  INTERCEPTION_MOUSE_BUTTON_4_DOWN = 0x40
  INTERCEPTION_MOUSE_BUTTON_4_UP = 0x80
  INTERCEPTION_MOUSE_BUTTON_5_DOWN = 0x100
  INTERCEPTION_MOUSE_BUTTON_5_UP = 0x200
  INTERCEPTION_MOUSE_WHEEL = 0x400
  INTERCEPTION_MOUSE_HWHEEL = 0x800

  # InterceptionMouseFlag
  INTERCEPTION_MOUSE_MOVE_RELATIVE      = 0x000
  INTERCEPTION_MOUSE_MOVE_ABSOLUTE      = 0x001
  INTERCEPTION_MOUSE_VIRTUAL_DESKTOP    = 0x002
  INTERCEPTION_MOUSE_ATTRIBUTES_CHANGED = 0x004
  INTERCEPTION_MOUSE_MOVE_NOCOALESCE    = 0x008
  INTERCEPTION_MOUSE_TERMSRV_SRC_SHADOW = 0x100

  # Screen wh
  SCREEN_WIDTH = 0
  SCREEN_HEIGHT = 0

  def __init__(self,interception,device):
    # wh取得
    self.SCREEN_WIDTH = user32.GetSystemMetrics(0)
    self.SCREEN_HEIGHT = user32.GetSystemMetrics(1)
    # interception インスタンス読み込み
    self.interception = interception
    self.device = device

  def get_absolute_x(self,x):
    return int((0xFFFF * x) / self.SCREEN_WIDTH)
  
  def get_absolute_y(self,y):
    return int((0xFFFF * y) / self.SCREEN_HEIGHT)

  def left_down(self,x,y,wait=0):
    time.sleep(wait/1000)
    stroke = _MOUSE_INPUT_DATA()
    stroke.Flags = self.INTERCEPTION_MOUSE_MOVE_ABSOLUTE
    stroke.ButtonFlags = self.INTERCEPTION_MOUSE_LEFT_BUTTON_DOWN
    stroke.LastX = self.get_absolute_x(x)
    stroke.LastY = self.get_absolute_y(y)
    self.interception.send(self.device,stroke)

  def left_up(self,x,y,wait=0):
    time.sleep(wait/1000)
    stroke = _MOUSE_INPUT_DATA()
    stroke.Flags = self.INTERCEPTION_MOUSE_MOVE_ABSOLUTE
    stroke.ButtonFlags = self.INTERCEPTION_MOUSE_LEFT_BUTTON_UP
    stroke.LastX = self.get_absolute_x(x)
    stroke.LastY = self.get_absolute_y(y)
    self.interception.send(self.device,stroke)

  def left_click(self,x,y,wait=0):
    time.sleep(wait/1000)
    self.left_down(x,y,0)
    time.sleep(0.05 * (0.8 + 0.4 * random()))
    self.left_up(x,y,0)

  def right_down(self,x,y,wait=0):
    time.sleep(wait/1000)
    stroke = _MOUSE_INPUT_DATA()
    stroke.Flags = self.INTERCEPTION_MOUSE_MOVE_ABSOLUTE
    stroke.ButtonFlags = self.INTERCEPTION_MOUSE_RIGHT_BUTTON_DOWN
    stroke.LastX = self.get_absolute_x(x)
    stroke.LastY = self.get_absolute_y(y)
    self.interception.send(self.device,stroke)

  def right_up(self,x,y,wait=0):
    time.sleep(wait/1000)
    stroke = _MOUSE_INPUT_DATA()
    stroke.Flags = self.INTERCEPTION_MOUSE_MOVE_ABSOLUTE
    stroke.ButtonFlags = self.INTERCEPTION_MOUSE_RIGHT_BUTTON_UP
    stroke.LastX = self.get_absolute_x(x)
    stroke.LastY = self.get_absolute_y(y)
    self.interception.send(self.device,stroke)

  def right_click(self,x,y,wait=0):
    time.sleep(wait/1000)
    self.right_down(x,y,0)
    time.sleep(0.05 * (0.8 + 0.4 * random()))
    self.right_up(x,y,0)

  def move_absolute(self,x,y,wait=0):
    time.sleep(wait/1000)
    stroke = _MOUSE_INPUT_DATA()
    stroke.Flags = self.INTERCEPTION_MOUSE_MOVE_ABSOLUTE
    stroke.LastX = self.get_absolute_x(x)
    stroke.LastY = self.get_absolute_y(y)
    self.interception.send(self.device,stroke)

  def move_relative(self,x,y,wait=0):
    time.sleep(wait/1000)
    stroke = _MOUSE_INPUT_DATA()
    stroke.Flags = self.INTERCEPTION_MOUSE_MOVE_RELATIVE
    stroke.LastX = x
    stroke.LastY = y
    self.interception.send(self.device,stroke)

  def wheel(self,scroll,wait=0):
    time.sleep(wait/1000)
    stroke = _MOUSE_INPUT_DATA()
    stroke.Flags = self.INTERCEPTION_MOUSE_MOVE_RELATIVE
    stroke.ButtonFlags = self.INTERCEPTION_MOUSE_WHEEL
    stroke.ButtonData = scroll
    stroke.LastX = 0
    stroke.LastY = 0
    self.interception.send(self.device,stroke)



##########################
# キーボードのSC
##########################

# https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes?redirectedfrom=MSDN
# 上のやつを全部MapVirtualKeyA関数で変換しただけ
SC_DECIMAL = {  'tab': 15,     # Special Keys
                'alt': 56,
                'space': 57,
                'back': 14,
                'lshift': 42,
                'shift': 42,
                'ctrl': 29,
                'del': 83,
                'end': 79,
                'pgup': 73,
                'pageup': 73,
                'pagedown': 81,
                'pgdown': 81,
                'insert': 82,
                'esc': 1,
                'enter': 28,
                'oem_plus': 39,
                'oem_minus': 12,
                'delete': 83,
                'home': 71,
                '^': 13,
                'e/j': 41,       # 半角/全角
                
                'left': 75,     # Arrow keys
                'up': 72,
                'right': 77,
                'down': 80,
                
                '0': 11,        # Numbers
                '1': 2,
                '2': 3,
                '3': 4,
                '4': 5,
                '5': 6,
                '6': 7,
                '7': 8,
                '8': 9,
                '9': 10,
                
                'f1': 59,
                'f2': 60,
                'f3': 61,
                'f4': 62,
                'f5': 63,
                'f6': 64,
                'f7': 65,
                'f8': 66,
                'f9': 67,
                'f10': 68,
                'f11': 87,
                'f12': 88,      # Function keys
                
                'a': 30,        # Letters
                'b': 48,
                'c': 46,
                'd': 32,
                'e': 18,
                'f': 33,
                'g': 34,
                'h': 35,
                'i': 23,
                'j': 36,
                'k': 37,
                'l': 38,
                'm': 50,
                'n': 49,
                'o': 24,
                'p': 25,
                'q': 16,
                'r': 19,
                's': 31,
                't': 20,
                'u': 22,
                'v': 47,
                'w': 17,
                'x': 45,
                'y': 21,
                'z': 44 }

##########################
# キーボード操作
##########################

class keyInput:
  # キーボード操作の定数
  INTERCEPTION_KEY_DOWN = 0x0
  INTERCEPTION_KEY_UP = 0x1
  INTERCEPTION_KEY_E0 = 0x2
  INTERCEPTION_KEY_E1 = 0x4
  INTERCEPTION_KEY_TERMSRV_SET_LED = 0x8
  INTERCEPTION_KEY_TERMSRV_SHADOW = 0x10
  INTERCEPTION_KEY_TERMSRV_VKPACKET = 0x20

  def __init__(self,interception,device):
    # interception インスタンス読み込み
    self.interception = interception
    self.device = device

  def click(self,key,wait=0):
    if (key in SC_DECIMAL):
      time.sleep(wait/1000)
      self.down(key,0)
      time.sleep(0.05 * (0.8 + 0.4 * random()))
      self.up(key,0)
      time.sleep(0.05 * (0.8 + 0.4 * random()))

  def down(self,key,wait=0):
    if (key in SC_DECIMAL):
      time.sleep(wait/1000)
      stroke = _KEYBOARD_INPUT_DATA()
      stroke.MakeCode = SC_DECIMAL[key]
      stroke.Flags = self.INTERCEPTION_KEY_DOWN
      self.interception.send(self.device,stroke)

  def up(self,key,wait=0):
    if (key in SC_DECIMAL):
      time.sleep(wait/1000)
      stroke = _KEYBOARD_INPUT_DATA()
      stroke.MakeCode = SC_DECIMAL[key]
      stroke.Flags = self.INTERCEPTION_KEY_UP
      self.interception.send(self.device,stroke)
