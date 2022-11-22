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


DEVPKEY_DeviceClass_IconPath = DEFINE_PROPERTYKEY(
    0x259abffc,
    0x50a7,
    0x47ce,
    0xaf,
    0x8,
    0x68,
    0xc9,
    0xa7,
    0xd7,
    0x33,
    0x66,
    12
)

DEVPKEY_AudioEndpointPlugin_FactoryCLSID = (
    DEFINE_PROPERTYKEY(
        0x12d83bd7,
        0xcf12,
        0x46be,
        0x85,
        0x40,
        0x81,
        0x27,
        0x10,
        0xd3,
        0x2,
        0x1c,
        1
    )
)

DEVPKEY_AudioEndpointPlugin_DataFlow = (
    DEFINE_PROPERTYKEY(
        0x12d83bd7,
        0xcf12,
        0x46be,
        0x85,
        0x40,
        0x81,
        0x27,
        0x10,
        0xd3,
        0x2,
        0x1c,
        2
    )
)

DEVPKEY_AudioEndpointPlugin_PnPInterface = (
    DEFINE_PROPERTYKEY(
        0x12d83bd7,
        0xcf12,
        0x46be,
        0x85,
        0x40,
        0x81,
        0x27,
        0x10,
        0xd3,
        0x2,
        0x1c,
        3
    )
)
