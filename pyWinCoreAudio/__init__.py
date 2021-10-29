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

__version__ = '0.1.0a'
__author__ = 'Kevin G. Schlosser'
__description__ = 'Python bindings to Microsoft Windows® Core Audio'
__url__ = 'https://github.com/kdschlosser/pyWinCoreAudio'



import comtypes as _comtypes
import weakref as _weakref

from .signal import (
    ON_DEVICE_ADDED,
    ON_DEVICE_REMOVED,
    ON_DEVICE_PROPERTY_CHANGED,
    ON_DEVICE_STATE_CHANGED,
    ON_ENDPOINT_VOLUME_CHANGED,
    ON_ENDPOINT_DEFAULT_CHANGED,
    ON_SESSION_VOLUME_DUCK,
    ON_SESSION_VOLUME_UNDUCK,
    ON_SESSION_CREATED,
    ON_SESSION_NAME_CHANGED,
    ON_SESSION_GROUPING_CHANGED,
    ON_SESSION_ICON_CHANGED,
    ON_SESSION_DISCONNECT,
    ON_SESSION_VOLUME_CHANGED,
    ON_SESSION_STATE_CHANGED,
    ON_SESSION_CHANNEL_VOLUME_CHANGED,
    ON_PART_CHANGE
)

_device_enumerator = None


def devices(message=True):
    if message:
        print(
            '****************** IMPORTANT ******************\n'
            'DO NOT hold a reference to any of the objects\n'
            'that are produced from this library. If you do\n'
            'there is a high probability of getting a memory\n'
            'leak.\n\n'
            'Make sure you call pyWinCoreAudio.stop() when\n'
            'this library is no longer going to be used.\n\n'
            'The object returned from pyWinCoreAudio.devices\n'
            'is a weak reference and this is the only object\n'
            'from this library that you are allowed to hold\n'
            'a reference to. The object returned is a \n'
            'callable and must be called each and every\n'
            'time before using it. DO NOT hold a reference\n'
            'to the object that is returned.\n\n'
            'The above is really important so that the\n'
            'library can function properly and no errors\n'
            'will take place.\n'
            '***********************************************\n'
        )

    global _device_enumerator

    if _device_enumerator is None:
        _comtypes.CoInitialize()
        from .mmdeviceapi import IMMDeviceEnumerator

        _device_enumerator = IMMDeviceEnumerator()

    return _weakref.ref(_device_enumerator)


def stop():
    global _device_enumerator

    if _device_enumerator is not None:
        _device_enumerator = None
        _comtypes.CoUninitialize()
