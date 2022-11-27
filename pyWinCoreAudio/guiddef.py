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

import ctypes.wintypes
import comtypes
import ctypes


# noinspection PyTypeChecker,PyCallingNonCallable
class GUID(comtypes.GUID):

    def __bool__(self):
        return self != GUID_NULL

    def __copy__(self):
        return GUID(str(self))

    def __init__(self, *guid):
        """a144ed38 - 8e12 - 4de4 - 9d96 -e6 47 40 b1 a5 24}
        Data1 = 0xa144ed38
        Data2 = 0x8e12
        Data3 = 0x4de4
        data4[0] = -0x9d
        data4[1] = 0x96
        data4[2] = 0xe6
        data4[3] = 0x47
        data4[4] = 0x40
        data4[5] = 0xb1
        data4[6] = 0xa5
        data4[7] = 0x24
        """

        if not guid:
            guid = None

        elif len(guid) == 1:
            guid = guid[0]
            try:
                if not guid.startswith('{'):
                    guid = '{' + guid + '}'
            except AttributeError:
                pass

        else:
            l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8 = guid

            w3 = b1 << 8 | b2

            ll = b3 << 40
            ll |= b4 << 32
            ll |= b5 << 24
            ll |= b6 << 16
            ll |= b7 << 8
            ll |= b8

            guid = [
                hex(l)[2:].upper().zfill(8),
                hex(w1)[2:].upper().zfill(4),
                hex(w2)[2:].upper().zfill(4),
                hex(w3)[2:].upper().zfill(4),
                hex(ll)[2:].upper().zfill(12)
            ]

            guid = '{' + '-'.join(guid) + '}'

        comtypes.GUID.__init__(self, guid)


GUID_NULL = GUID()


def InlineIsEqualGUID(rguid1, rguid2):
    return rguid1 == rguid2


def IsEqualGUID(rguid1, rguid2):
    return InlineIsEqualGUID(rguid1, rguid2)


# noinspection PyPep8
def DEFINE_GUID(l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8):
    return GUID(l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8)


# noinspection PyPep8
def DEFINE_OLEGUID(l, w1, w2):
    return DEFINE_GUID(l, w1, w2, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46)


# noinspection PyPep8
def EXTERN_GUID(l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8):
    return GUID(l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8)


# noinspection PyPep8
def DEFINE_GUIDEX(l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8):
    return GUID(l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8)


def MIDL_INTERFACE(g):
    return GUID(g)


LPGUID = ctypes.POINTER(GUID)
LPCGUID = ctypes.POINTER(GUID)


IID = GUID
LPIID = ctypes.POINTER(IID)
IID_NULL = GUID()


def IsEqualIID(riid1, riid2):
    return IsEqualGUID(riid1, riid2)


CLSID = GUID
LPCLSID = ctypes.POINTER(CLSID)
CLSID_NULL = GUID()


def IsEqualCLSID(rclsid1, rclsid2):
    return IsEqualGUID(rclsid1, rclsid2)


FMTID = GUID
LPFMTID = ctypes.POINTER(FMTID)
FMTID_NULL = GUID()


def IsEqualFMTID(rfmtid1, rfmtid2):
    return IsEqualGUID(rfmtid1, rfmtid2)


REFGUID = GUID
REFIID = IID
REFCLSID = IID
REFFMTID = IID


def DEFINE_GUIDSTRUCT(guid):
    return GUID(guid)


class NodeTypeGUID(GUID):
    _instances = {}

    @classmethod
    def get(cls, guid):
        return cls._instances.get(guid, None)

    def __init__(self, name, *guid):
        super(NodeTypeGUID, self).__init__(*guid)
        self._instances[guid] = self
        self._label = name

    @property
    def label(self):
        return self._label

    def __str__(self):
        return self._label


def DEFINE_USB_TERMINAL_GUID(*hex_code):
    return (
        0xDFF219E0 + hex_code[0],
        0xF70F,
        0x11D0,
        0xB9,
        0x17,
        0x00,
        0xA0,
        0xC9,
        0x22,
        0x31,
        0x96
    )


def DEFINE_GUIDNAMED(guid):
    if isinstance(guid, GUID):
        return guid
    return GUID(guid)


def DECLSPEC_UUID(guid):
    return GUID(guid)


class PropertyKeyMeta(type(ctypes.Structure)):
    _keys = {}

    def __init__(cls, name, bases, dct):
        super(PropertyKeyMeta, cls).__init__(name, bases, dct)
        setattr(cls, '_keys', PropertyKeyMeta._keys)

    def __call__(cls, *args, **kwargs):
        self = super(PropertyKeyMeta, cls).__call__(*args, **kwargs)
        key = (str(self.fmtid), self.pid)

        if key == ('{00000000-0000-0000-0000-000000000000}', 0):
            return self

        if key not in PropertyKeyMeta._keys:
            PropertyKeyMeta._keys[key] = self

        return PropertyKeyMeta._keys[key]


class _tagPROPERTYKEY(ctypes.Structure, metaclass=PropertyKeyMeta):
    _keys = {}

    _fields_ = [
        ('fmtid', GUID),
        ('pid', ctypes.wintypes.ULONG),
    ]

    def __eq__(self, other):
        if not isinstance(other, _tagPROPERTYKEY):
            return False

        return (
            str(other.fmtid) == str(self.fmtid) and
            other.pid == self.pid
        )

    def __ne__(self, other):
        return not self.__eq__(other)

    def __str__(self):
        key = (str(self.fmtid), self.pid)

        if key in self._keys:
            return str(self._keys[key])

        return 'PROPERTYKEY(' + str(self.fmtid) + ', ' + str(self.pid) + ')'

    def get(self, endpoint):
        key = (str(self.fmtid), self.pid)

        if key in self._keys:
            key = self._keys[key]
        else:
            key = self

        data = endpoint.get_property(key)
        data = key.decode(endpoint, data)
        return data

    def decode(self, endpoint, data):
        return data


PROPERTYKEY = _tagPROPERTYKEY
PPROPERTYKEY = ctypes.POINTER(PROPERTYKEY)


def DEFINE_PROPERTYKEY(l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8, key):
    return PROPERTYKEY(GUID(l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8), key)


uuid = GUID



