"""
interceptionのドライバ処理さくっと
"""

from ctypes import *

kernel32 = cdll.LoadLibrary("Kernel32.dll")

class interception:
  def __init__(self):
    self.devices_handle = [0] * 20
    self.devices_unempty = [0] * 20
    self.MAX_DEVICE = 20
    self.MAX_KEYBOARD = 10
    self.MAX_MOUSE = 10
    self.load()

  def load(self):
    """
    driverのhandleを一覧取得
    """
    device_base = "\\\\.\\interception"
    for  i in range(self.MAX_DEVICE):
      device_name = device_base + str(i).zfill(2)
      self.devices_handle[i] = kernel32.CreateFileA(device_name.encode("ascii"),0x80000000, 0, 0, 3, 0, 0)
      self.devices_unempty[i] = kernel32.CreateEventA(None, True, False, None)

  def send(self,device,stroke):
    if (self.is_invalid(device) or self.devices_handle[device-1] == 0):
      return 0

    # @TODO 本当はoutを参照渡しして結果を取得したいけど、そうすると止まるから保留
    out = c_int32()
    res = kernel32.DeviceIoControl(self.devices_handle[device-1],0x222080,byref(stroke) ,sizeof(stroke),out,sizeof(out),None,None)
    return out


  def get_hwid_list(self):
    hwid_list = [''] * 20
    hw_count = 0
    for i in range(self.MAX_DEVICE):
      tmp_hwid = self.get_hwid(self.devices_handle[i])
      if (tmp_hwid != ''):
        hwid_list[hw_count] = tmp_hwid
        hw_count += 1
    return hwid_list[0:hw_count]
  
  def search_device_handle(self,hwid):
    for i in range(self.MAX_DEVICE):
      device_hwid = self.get_hwid(self.devices_handle[i])
      if (device_hwid.find(hwid) != -1):
        return i + 1
    return 0

  def get_hwid(self,handle):
    out = create_unicode_buffer(500)
    out_size = c_int32(0)
    kernel32.DeviceIoControl(handle,0x222200,None,0,out,500,byref(out_size),None)
    return out.value

  def is_keyboard(self,device):
    return device + 1 > 0 and device + 1 <= self.MAX_KEYBOARD

  def is_mouse(self,device):
    return device + 1 > self.MAX_KEYBOARD and device + 1 <= self.MAX_KEYBOARD + self.MAX_MOUSE

  def is_invalid(self,device):
    return device + 1 <= 0 or device + 1 > (self.MAX_KEYBOARD + self.MAX_MOUSE)
