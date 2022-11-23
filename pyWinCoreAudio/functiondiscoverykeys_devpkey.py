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

# from .data_types import *

from .guiddef import PROPERTYKEY as _PROPERTYKEY, GUID


class PK(_PROPERTYKEY):
    def __str__(self):
        for name, pkey in globals().items():
            if not name.startswith('PKEY_'):
                continue

            if (
                str(pkey.fmtid) == str(self.fmtid) and
                pkey.pid == self.pid
            ):
                return name

        return _PROPERTYKEY.__str__(self)


def DEFINE_PROPERTYKEY(l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8, key):
    return PK(GUID(l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8), key)


PKEY_Device_FriendlyName = DEFINE_PROPERTYKEY(
    0xA45C254E,
    0xDF1C,
    0x4EFD,
    0x80,
    0x20,
    0x67,
    0xD1,
    0x46,
    0xA8,
    0x50,
    0xE0,
    14
)

PKEY_Device_DeviceDesc = DEFINE_PROPERTYKEY(
    0xA45C254E,
    0xDF1C,
    0x4EFD,
    0x80,
    0x20,
    0x67,
    0xD1,
    0x46,
    0xA8,
    0x50,
    0xE0,
    2
)

PKEY_DeviceInterface_FriendlyName = DEFINE_PROPERTYKEY(
    0x026E516E,
    0xB814,
    0x414B,
    0x83,
    0xCD,
    0x85,
    0x6D,
    0x6F,
    0xEF,
    0x48,
    0x22,
    2
)
