# -*- coding: utf-8 -*-
#
# This file is part of EventGhost.
# Copyright © 2005-2016 EventGhost Project <http://www.eventghost.net/>
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

from .mmdeviceapi import PIMMDevice
from .data_types import *

import ctypes


class AudioExtensionParams(ctypes.Structure):
    _fields_ = [
        ('AddPageParam', LPARAM),
        ('pEndpoint', PIMMDevice),
        ('pPnpInterface', PIMMDevice),
        ('pPnpDevnode', PIMMDevice),

    ]


class tagDIRECTX_AUDIO_ACTIVATION_PARAMS(ctypes.Structure):
    _fields_ = [
        ('cbDirectXAudioActivationParams', DWORD),
        ('guidAudioSession', GUID),
        ('dwAudioStreamFlags', DWORD),
    ]


DIRECTX_AUDIO_ACTIVATION_PARAMS = tagDIRECTX_AUDIO_ACTIVATION_PARAMS
PDIRECTX_AUDIO_ACTIVATION_PARAMS = POINTER(DIRECTX_AUDIO_ACTIVATION_PARAMS)


class KSPROPERTY(ctypes.Structure):
    _fields_ = [
        ('Set', GUID),
        ('Id', ULONG),
        ('Flags', ULONG)
    ]


PKSPROPERTY = POINTER(KSPROPERTY)


class KSMETHOD(ctypes.Structure):
    _fields_ = [
        ('Set', GUID),
        ('Id', ULONG),
        ('Flags', ULONG)
    ]


PKSMETHOD = POINTER(KSMETHOD)


class KSEVENT(ctypes.Structure):
    _fields_ = [
        ('Set', GUID),
        ('Id', ULONG),
        ('Flags', ULONG)
    ]


PKSEVENT = POINTER(KSEVENT)


class KSIDENTIFIER_STRUCTURE(ctypes.Structure):
    _fields_ = [
        ('Set', GUID),
        ('Id', ULONG),
        ('Flags', ULONG),
    ]


class KSIDENTIFIER_UNION(ctypes.Union):
    _anonymous_ = ('struct',)
    _fields_ = [
        ('struct', KSIDENTIFIER_STRUCTURE)
    ]


PKSIDENTIFIER_STRUCTURE = POINTER(KSIDENTIFIER_STRUCTURE)


class KSIDENTIFIER(ctypes.Structure):
    _anonymous_ = ('union',)
    _fields_ = [
        ('Alignment', LONGLONG),
        ('union', KSIDENTIFIER_UNION)
    ]


PKSIDENTIFIER = POINTER(KSIDENTIFIER)
