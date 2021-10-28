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

from .data_types import *
import ctypes
import comtypes
from .mmdeviceapi import ERole
from .audioclient import PWAVEFORMATEX
from .propertystore import (
    PPROPERTYKEY,
    PPROPVARIANT
)


IID_IPolicyConfig = IID(
    '{f8679f50-850a-41cf-9c72-430f290290c8}'
)
IID_IPolicyConfigVista = IID(
    '{568b9108-44bf-40b4-9006-86afe5b5a620}'
)
IID_AudioSes = (
    '{00000000-0000-0000-0000-000000000000}'
)
CLSID_PolicyConfigClient = IID(
    '{870af99c-171d-4f9e-af0d-e63df40c2bc9}'
)
CLSID_PolicyConfigVistaClient = IID(
    '{294935CE-F637-4E7C-A41B-AB255460B862}'
)


class DeviceSharedMode(ctypes.Structure):
    _fields_ = [
        ('dummy_', INT)
    ]


PDeviceSharedMode = POINTER(DeviceSharedMode)


class IPolicyConfig(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IPolicyConfig
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetMixFormat',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['out'], POINTER(PWAVEFORMATEX), 'pFormat')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetDeviceFormat',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], BOOL, 'bDefault'),
            (['out'], POINTER(PWAVEFORMATEX), 'pFormat')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'ResetDeviceFormat',
            (['in'], LPCWSTR, 'pwstrDeviceId')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetDeviceFormat',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], PWAVEFORMATEX, 'pEndpointFormat'),
            (['in'], PWAVEFORMATEX, 'pMixFormat')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetProcessingPeriod',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], BOOL, 'bDefault'),
            (['out'], LPREFERENCE_TIME, 'hnsDefaultDevicePeriod'),
            (['out'], LPREFERENCE_TIME, 'hnsMinimumDevicePeriod')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetProcessingPeriod',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], LPREFERENCE_TIME, 'hnsDevicePeriod')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetShareMode',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['out'], PDeviceSharedMode, 'pMode')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetShareMode',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], PDeviceSharedMode, 'pMode')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetPropertyValue',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], PPROPERTYKEY, 'key'),
            (['out'], PPROPVARIANT, 'pValue')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetPropertyValue',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], PPROPERTYKEY, 'key'),
            (['in'], PPROPVARIANT, 'pValue')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetDefaultEndpoint',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], ERole, 'ERole')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetEndpointVisibility',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], BOOL, 'bVisible')
        )
    )


# noinspection PyTypeChecker
PIPolicyConfig = POINTER(IPolicyConfig)


class IPolicyConfigVista(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IPolicyConfigVista
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetMixFormat',
            (['in'], LPCWSTR, 'wszDeviceId'),
            (['out'], POINTER(PWAVEFORMATEX), 'pFormat')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetDeviceFormat',
            (['in'], LPCWSTR, 'wszDeviceId'),
            (['in'], BOOL, 'bDefault'),
            (['out'], POINTER(PWAVEFORMATEX), 'pFormat')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetDeviceFormat',
            (['in'], LPCWSTR, 'wszDeviceId'),
            (['in'], PWAVEFORMATEX, 'pEndpointFormat'),
            (['in'], PWAVEFORMATEX, 'pMixFormat')
        ),

        COMMETHOD(
            [],
            HRESULT,
            'GetProcessingPeriod',
            (['in'], LPCWSTR, 'wszDeviceId'),
            (['in'], BOOL, 'bDefault'),
            (['out'], LPREFERENCE_TIME, 'hnsDefaultDevicePeriod'),
            (['out'], LPREFERENCE_TIME, 'hnsMinimumDevicePeriod')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetProcessingPeriod',
            (['in'], LPCWSTR, 'wszDeviceId'),
            (['in'], LPREFERENCE_TIME, 'hnsDevicePeriod')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetShareMode',
            (['in'], LPCWSTR, 'wszDeviceId'),
            (['out'], PDeviceSharedMode, 'pMode')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetShareMode',
            (['in'], LPCWSTR, 'wszDeviceId'),
            (['in'], PDeviceSharedMode, 'pMode')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetPropertyValue',
            (['in'], LPCWSTR, 'wszDeviceId'),
            (['in'], PPROPERTYKEY, 'key'),
            (['out'], PPROPVARIANT, 'pValue')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetPropertyValue',
            (['in'], LPCWSTR, 'wszDeviceId'),
            (['in'], PPROPERTYKEY, 'key'),
            (['in'], PPROPVARIANT, 'pValue')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetDefaultEndpoint',
            (['in'], LPCWSTR, 'wszDeviceId'),
            (['in'], ERole, 'ERole')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetEndpointVisibility',
            (['in'], LPCWSTR, 'wszDeviceId'),
            (['in'], BOOL, 'bVisible')
        )
    )


# noinspection PyTypeChecker
PIPolicyConfigVista = POINTER(IPolicyConfigVista)


class AudioSes(object):
    name = u'AudioSes'
    _reg_typelib_ = (IID_AudioSes, 1, 0)


class CPolicyConfigClient (comtypes.CoClass):
    _reg_clsid_ = CLSID_PolicyConfigClient
    _idlflags_ = []
    _reg_typelib_ = (IID_AudioSes, 1, 0)
    _com_interfaces_ = [IPolicyConfig]


class CPolicyConfigVistaClient(comtypes.CoClass):
    _reg_clsid_ = CLSID_PolicyConfigVistaClient
    _idlflags_ = []
    _reg_typelib_ = (IID_AudioSes, 1, 0)
    _com_interfaces_ = [IPolicyConfigVista]
