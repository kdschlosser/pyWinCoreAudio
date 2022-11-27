# -*- coding: utf-8 -*-
#
# This file is part of EventGhost.
# Copyright © 2005-2021 EventGhost Project <http://www.eventghost.net/>
#
# EventGhost is free software: you can redistribute it and/or modify it under
# the terms of the GNU General Public License as published by the Free
# Software Foundation, either version 2 of the License, or (at your option)
# any later version.
#
# EventGhost is distributed in the hope that it will be useful, but WITHOUT
# ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
# FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
# more details.
#
# You should have received a copy of the GNU General Public License along
# with EventGhost. If not, see <http://www.gnu.org/licenses/>.

from .data_types import *
import threading
import ctypes
import comtypes
import os
from io import BytesIO
import winreg


RT_CURSOR = 1
RT_BITMAP = 2
RT_ICON = 3
RT_MENU = 4
RT_DIALOG = 5
RT_STRING = 6
RT_FONTDIR = 7
RT_FONT = 8
RT_ACCELERATOR = 9
RT_RCDATA = 10
RT_MESSAGETABLE = 11
DIFFERENCE = 11
RT_GROUP_CURSOR = (RT_CURSOR + DIFFERENCE)
RT_GROUP_ICON = (RT_ICON + DIFFERENCE)
RT_VERSION = 16
RT_DLGINCLUDE = 17
RT_PLUGPLAY = 19
RT_VXD = 20
RT_ANICURSOR = 21
RT_ANIICON = 22
RT_HTML = 23


def convert_to_string(data, length=None):
    if not isinstance(data, str):
        count = 0

        chars = []
        try:
            while True:
                char = data[count]
                if length is None:
                    if isinstance(char, int) and char == 0x00:
                        dta = bytearray(chars).decode('utf-8')
                        break
                    elif isinstance(char, str) and char == '\x00':
                        dta = ''.join(chars)
                        break
                elif length == count:
                    chars += [char]

                    if isinstance(char, int):
                        dta = bytearray(chars).decode('utf-8')
                        break
                    else:
                        dta = ''.join(chars)
                        break

                chars += [char]
                count += 1

            data = dta
        except ValueError:
            return ''

    return data


def run_in_thread(func, *args, **kwargs):
    t = threading.Thread(target=func, args=args, kwargs=kwargs)
    t.daemon = True
    t.start()
    return t


def convert_triplet_to_rgb(triplet):
    if not triplet:
        return 0, 0, 0

    r, g, b = bytearray.fromhex(hex(triplet)[2:].replace('L', '').zfill(6))
    return r, g, b


icons = {}

kernel32 = ctypes.windll.kernel32

libc = ctypes.cdll.msvcrt
libc.memcpy.argtypes = [LPVOID, LPVOID, SIZE_T]
libc.memcpy.restype = LPCSTR


def get_icon(icon):
    global icons

    if not isinstance(icon, str):
        icn = ''
        for item in icon:
            icn += item

        icon = icn

    if icon in icons:
        return icons[icon]

    try:
        icon_path, icon_name = icon.replace('@', '').split(',-')
        icon_name = int(icon_name)
    except ValueError:
        icon_path = icon.replace('@', '')
        icon_name = 1

    try:
        hlib = kernel32.LoadLibraryExW(
            os.path.expandvars(icon_path),
            None,
            0x00000020
        )

        # This part almost identical to C++
        hResInfo = ctypes.windll.kernel32.FindResourceW(
            hlib,
            icon_name,
            RT_ICON
        )
        size = ctypes.windll.kernel32.SizeofResource(
            hlib,
            hResInfo
        )
        rec = kernel32.LoadResource(hlib, hResInfo)
        mem_pointer = kernel32.LockResource(rec)

        # And this is some differ (copy data to Python buffer)
        binary_data = (ctypes.c_ubyte * size)()
        libc.memcpy(binary_data, mem_pointer, size)

        f = BytesIO()
        f.write(bytearray(binary_data))
        f.seek(0)
        icons[icon] = f
        return f

    except comtypes.COMError:
        import traceback
        traceback.print_exc()


def remap(value, old_min, old_max, new_min, new_max):
    if (
            isinstance(value, float) or
            isinstance(old_min, float) or
            isinstance(old_max, float) or
            isinstance(new_min, float) or
            isinstance(new_max, float)
    ):
        type_ = float

    else:
        type_ = int

    value = type_(value)
    old_min = type_(old_min)
    old_max = type_(old_max)
    new_max = type_(new_max)
    new_min = type_(new_min)

    old_range = old_max - old_min  # type: ignore
    new_range = new_max - new_min  # type: ignore

    if old_range == 0:
        raise ValueError('Input range ({}-{}) is empty'.format(
            old_min, old_max))

    if new_range == 0:
        raise ValueError('Output range ({}-{}) is empty'.format(
            new_min, new_max))

    new_value = (value - old_min) * new_range  # type: ignore

    if type_ == int:
        new_value //= old_range  # type: ignore
    else:
        new_value /= old_range  # type: ignore

    new_value += new_min  # type: ignore

    return new_value


# DWORD GetProcessImageFileNameW(
#   [in]  HANDLE hProcess,
#   [out] LPWSTR lpImageFileName,
#   [in]  DWORD  nSize
# );

_kernel32 = ctypes.windll.Kernel32
_psapi = ctypes.windll.Psapi

_GetProcessImageFileNameW = _psapi.GetProcessImageFileNameW
_GetProcessImageFileNameW.rstype = DWORD

# HANDLE OpenProcess(
#   [in] DWORD dwDesiredAccess,
#   [in] BOOL  bInheritHandle,
#   [in] DWORD dwProcessId
# );
_OpenProcess = _kernel32.OpenProcess
_OpenProcess.restype = HANDLE

# BOOL CloseHandle(
#   [in] HANDLE hObject
# );

_CloseHandle = _kernel32.CloseHandle
_CloseHandle.restype = BOOL

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000


def get_process_name(id_):
    hprocess = _OpenProcess(
        DWORD(PROCESS_QUERY_LIMITED_INFORMATION),
        BOOL(1),
        DWORD(id_)
    )

    buf = ctypes.create_unicode_buffer(255)

    _GetProcessImageFileNameW(
        HANDLE(hprocess),
        buf,
        DWORD(255)
    )

    _CloseHandle(HANDLE(hprocess))

    value = buf.value

    if value:
        return os.path.split(os.path.splitext(value)[0])[-1]



import struct

class Registry(object):

    def __init__(self, key):
        self.key = key

    def __iter__(self):
        num_keys = winreg.QueryInfoKey(self.key)[0]
        for i in range(0, num_keys):
            key_name = winreg.EnumKey(self.key, i)
            key = winreg.OpenKey(self.key, key_name)
            yield Registry(key)

    @property
    def values(self):
        num_values = winreg.QueryInfoKey(self.key)[1]
        for i in range(0, num_values):
            name, data, d_type = winreg.EnumValue(self.key, i)
            winreg.REG_BINARY
            winreg.REG_DWORD
            winreg.REG_DWORD_LITTLE_ENDIAN
            winreg.REG_DWORD_BIG_ENDIAN

            winreg.REG_MULTI_SZ
            winreg.REG_NONE
            winreg.REG_QWORD
            winreg.REG_QWORD_LITTLE_ENDIAN
            winreg.REG_SZ


class ValueBase(object):

    def __init__(self, key, name, data):
        self.key = key
        self.name = name
        self.data = data


class ValueBinary(ValueBase):

    def __iter__(self):
        unpacked = struct.unpack('<Q', self.data)
        unpacked.bit_length()

        _winreg.SetValueEx(key, "Attributes", 0, _winreg.REG_BINARY, s)









import errno, os, winreg


def find_reg_key(root_key, key_name):
    for arch_key in (winreg.KEY_WOW64_32KEY, winreg.KEY_WOW64_64KEY):
        root_key


        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", 0, winreg.KEY_READ | arch_key)


def find_reg_value(root, path, value_name):
    for arch_key in (winreg.KEY_WOW64_32KEY, winreg.KEY_WOW64_64KEY):
        key = winreg.OpenKey(root, path, 0, winreg.KEY_READ | arch_key)
        num_keys, num_values = winreg.QueryInfoKey(key)[:2]
        for i in range(0, num_values):
            skey_name = winreg.EnumKey(key, i)
            skey = winreg.OpenKey(key, skey_name)
            try:
                print(winreg.QueryValueEx(skey, 'DisplayName')[0])
            except OSError as e:
                if e.errno == errno.ENOENT:
                    # DisplayName doesn't exist in this skey
                    pass
            finally:
                skey.Close()

