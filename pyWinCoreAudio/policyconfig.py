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

import platform
from .data_types import *
import ctypes
import comtypes
# from . import utils
from .mmdeviceapi import ERole, EDataFlow
from .inspectable import IInspectable
from .audioclient import PWAVEFORMATEX
from .propidl import PPROPVARIANT


WINVISTA = 6600
WINVISTA_SP1 = 6601
WINVISTA_SP2 = 6602
WIN7 = 7600
WIN7_SP1 = 7601
WIN8 = 9200
WIN81 = 9600
WIN10_THRESHOLD1 = 10240
WIN10_THRESHOLD2 = 10586
WIN10_REDSTONE1 = 14393
WIN10_REDSTONE3 = 16299
WIN10_21H2 = 19044

_major, _, _build = tuple(
    int(item) for item in platform.version().split('.')
)


if _build <= WINVISTA_SP2:
    IID_IPolicyConfig = IID(
        '{568B9108-44BF-40B4-9006-86AFE5b5A620}'
    )
elif _build <= WIN81:
    IID_IPolicyConfig = IID(
        '{F8679F50-850A-41CF-9C72-430F290290C8}'
    )
elif _build == WIN10_THRESHOLD1:
    IID_IPolicyConfig = IID(
        'CA286FC3-91FD-42C3-8E9B-CAAFA66242E3'
    )
elif _build == WIN10_THRESHOLD2:
    # Windows 10 Threshold 2
    IID_IPolicyConfig = IID(
        '{6BE54BE8-A068-4875-A49D-0C2966473B11}'
    )
else:
    IID_IPolicyConfig = IID(
        '{F8679F50-850A-41CF-9C72-430F290290C8}'
    )

if _major == 10 and WIN10_REDSTONE3 <= _build < WIN10_21H2:
    # '870AF99C-171D-4F9E-AF0D-E63DF40C2BC9'
    IID_IAudioPolicyConfig = IID(
        "{2A59116D-6C4F-45E0-A74F-707E3FEF9258}"
    )
elif _major == 10 and _build >= WIN10_21H2:
    _combase = ctypes.windll.combase
    _RoGetActivationFactory = _combase.RoGetActivationFactory
    _RoGetActivationFactory.restype = HRESULT

    IID_IAudioPolicyConfig = IID(
        "{AB3D4648-E242-459F-B02F-541C70306324}"
    )
else:
    IID_IAudioPolicyConfig = IID(
        "{00000000-0000-0000-0000-000000000000}"
    )

if _build <= WINVISTA_SP2:
    CLSID_PolicyConfigClient = IID(
        '{294935CE-F637-4E7C-A41B-AB255460B862}'
    )
else:
    CLSID_PolicyConfigClient = IID(
        '{870Af99C-171D-4f9E-AF0D-E63DF40C2BC9}'
    )


IID_AudioSes = (
    '{00000000-0000-0000-0000-000000000000}'
)


class IAudioPolicyConfig(IInspectable):
    _case_insensitive_ = False
    _iid_ = IID_IAudioPolicyConfig

    if _major == 10 and WIN10_REDSTONE3 <= _build < WIN10_21H2:
        _methods_ = (
            comtypes.STDMETHOD(
                HRESULT,
                'SetPersistedDefaultAudioEndpoint',
                [UINT, EDataFlow, ERole, HSTRING]
            ),
            comtypes.STDMETHOD(
                HRESULT,
                'GetPersistedDefaultAudioEndpoint',
                [UINT, EDataFlow, ERole, POINTER(HSTRING)]
            ),
            comtypes.STDMETHOD(
                HRESULT,
                'ClearAllPersistedApplicationDefaultEndpoints'
            )
        )

    elif _major == 10 and _build >= WIN10_21H2:
        _methods_ = (
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__add_CtxVolumeChange'
            ),
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__remove_CtxVolumeChanged'
            ),
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__add_RingerVibrateStateChanged'
            ),
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__remove_RingerVibrateStateChange'
            ),
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__SetVolumeGroupGainForId'
            ),
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__GetVolumeGroupGainForId'
            ),
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__GetActiveVolumeGroupForEndpointId'
            ),
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__GetVolumeGroupsForEndpoint'
            ),
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__GetCurrentVolumeContext'
            ),
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__SetVolumeGroupMuteForId'
            ),
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__GetVolumeGroupMuteForId'
            ),
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__SetRingerVibrateState'
            ),
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__GetRingerVibrateState'
            ),
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__SetPreferredChatApplication'
            ),
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__ResetPreferredChatApplication'
            ),
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__GetPreferredChatApplication'
            ),
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__GetCurrentChatApplications'
            ),
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__add_ChatContextChanged'
            ),
            COMMETHOD(
                [],
                HRESULT,
                '__incomplete__remove_ChatContextChanged'
            ),
            comtypes.STDMETHOD(
                HRESULT,
                'SetPersistedDefaultAudioEndpoint',
                [UINT, EDataFlow, ERole, HSTRING]
            ),
            comtypes.STDMETHOD(
                HRESULT,
                'GetPersistedDefaultAudioEndpoint',
                [UINT, EDataFlow, ERole, POINTER(HSTRING)]
            ),
            comtypes.STDMETHOD(
                HRESULT,
                'ClearAllPersistedApplicationDefaultEndpoints'
            )
        )
    else:
        _methods_ = ()

        @staticmethod
        def SetPersistedDefaultAudioEndpoint(*_, **__):
            pass

        @staticmethod
        def GetPersistedDefaultAudioEndpoint(*_, **__):
            pass

        @staticmethod
        def ClearAllPersistedApplicationDefaultEndpoints():
            pass


class DeviceSharedMode(ctypes.Structure):
    _fields_ = [
        ('dummy_', INT)
    ]


PDeviceSharedMode = POINTER(DeviceSharedMode)


class IPolicyConfig(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IPolicyConfig

    if _build <= WINVISTA_SP2:
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

        def ResetDeviceFormat(self, _):
            pass

    else:
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


DEVINTERFACE_AUDIO_RENDER = "{E6327CAD-DCEC-4949-AE8A-991E976A79D2}"
DEVINTERFACE_AUDIO_CAPTURE = "{2EEF81BE-33FA-4800-9670-1CD474972C3F}"
MMDEVAPI_TOKEN = "\\\\?\\SWD#MMDEVAPI#"


def logger(func):

    def wrapper(*args, **kwargs):

        # print('DEBUG: ', func.__name__)
        return func(*args, **kwargs)

    return wrapper


class PolicyConfigClient(object):

    @property
    @logger
    def _policy_config(self):
        if _major == 10 and _build >= WIN10_REDSTONE3:
            hClassName = HSTRING.from_string(
                "Windows.Media.Internal.AudioPolicyConfig"
            )
            policy_config = POINTER(POINTER(LPVOID))()

            _RoGetActivationFactory(
                hClassName,
                ctypes.byref(IID_IAudioPolicyConfig),
                ctypes.byref(policy_config)
            )

            policy_config = ctypes.cast(
                policy_config,
                POINTER(IAudioPolicyConfig)
            )

        else:
            policy_config = IAudioPolicyConfig()

        return policy_config

    @staticmethod
    @logger
    def _GenerateDeviceId(deviceId, flow):
        template = "{MMDEVAPI_TOKEN}{deviceId}#{flow}"
        if flow == EDataFlow.eRender:
            flow = DEVINTERFACE_AUDIO_RENDER
        else:
            flow = DEVINTERFACE_AUDIO_CAPTURE

        device_id = template.format(
            MMDEVAPI_TOKEN=MMDEVAPI_TOKEN,
            deviceId=deviceId,
            flow=flow
        )

        string = HSTRING.from_string(device_id)
        return string

    @staticmethod
    @logger
    def _UnpackDeviceId(deviceId):
        device_id = str(deviceId)

        if not device_id:
            return None

        device_id = device_id.replace(MMDEVAPI_TOKEN, '', 1)
        device_id = device_id.rsplit('#', 1)[0]

        return device_id

    @logger
    def SetSessionDefaultEndPoint(self, deviceId, flow, role, processId):
        if not deviceId:
            return

        deviceIdStr = self._GenerateDeviceId(deviceId, flow)
        self._policy_config.SetPersistedDefaultAudioEndpoint(
            processId,
            flow,
            role,
            deviceIdStr
        )

    @logger
    def GetSessionDefaultEndPoint(self, flow, role, processId):
        deviceId = HSTRING()

        self._policy_config.GetPersistedDefaultAudioEndpoint(
            processId,
            flow,
            role,
            ctypes.byref(deviceId)
        )

        unpacked = self._UnpackDeviceId(deviceId)

        return unpacked

    @logger
    def ResetAllSessionEndpoints(self):
        try:
            self._policy_config.ClearAllPersistedApplicationDefaultEndpoints()
        except:  # NOQA
            pass


class AudioSes(comtypes.CoClass):
    name = u'AudioSes'
    _reg_typelib_ = (IID_AudioSes, 1, 0)


class CPolicyConfigClient(comtypes.CoClass):
    _reg_clsid_ = CLSID_PolicyConfigClient
    _idlflags_ = []
    _reg_typelib_ = (IID_AudioSes, 1, 0)
    _com_interfaces_ = [IPolicyConfig]
