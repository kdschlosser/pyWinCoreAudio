
from .mmdeviceapi import *
from .propsys import IPropertyStore
from .AudioAPOTypes import *
from .audiomediatype import *
from .data_types import *
from .guiddef import PROPERTYKEY as _PROPERTYKEY
from .ksmedia import KSNODETYPE

import ctypes


def _HRESULT_TYPEDEF_(sc):
    return sc

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


# The object has already been initialized.
APOERR_ALREADY_INITIALIZED = _HRESULT_TYPEDEF_(0x887D0001)
# Object/structure is not initialized.
APOERR_NOT_INITIALIZED = _HRESULT_TYPEDEF_(0x887D0002)
# a pin supporting the format cannot be found.
APOERR_FORMAT_NOT_SUPPORTED = _HRESULT_TYPEDEF_(0x887D0003)
# Invalid CLSID in an APO Initialization structure
APOERR_INVALID_APO_CLSID = _HRESULT_TYPEDEF_(0x887D0004)
# Buffers overlap on an APO that does not accept in-place buffers.
APOERR_BUFFERS_OVERLAP = _HRESULT_TYPEDEF_(0x887D0005)
# APO is already in an unlocked state
APOERR_ALREADY_UNLOCKED = _HRESULT_TYPEDEF_(0x887D0006)
# number of input or output connections not valid on this APO
APOERR_NUM_CONNECTIONS_INVALID = _HRESULT_TYPEDEF_(0x887D0007)
# Output maxFrameCount not large enough.
APOERR_INVALID_OUTPUT_MAXFRAMECOUNT = _HRESULT_TYPEDEF_(0x887D0008)
# Invalid connection format for this operation
APOERR_INVALID_CONNECTION_FORMAT = _HRESULT_TYPEDEF_(0x887D0009)
# APO is locked ready to process and can not be changed
APOERR_APO_LOCKED = _HRESULT_TYPEDEF_(0x887D000A)
# Invalid coefficient count
APOERR_INVALID_COEFFCOUNT = _HRESULT_TYPEDEF_(0x887D000B)
# Invalid coefficient
APOERR_INVALID_COEFFICIENT = _HRESULT_TYPEDEF_(0x887D000C)
# an invalid curve parameter was specified
APOERR_INVALID_CURVE_PARAM = _HRESULT_TYPEDEF_(0x887D000D)
# Invalid auxiliary input index
APOERR_INVALID_INPUTID = _HRESULT_TYPEDEF_(0x887D000E)

# Signatures for data structures.
APO_CONNECTION_DESCRIPTOR_SIGNATURE = 'ACDS'
APO_CONNECTION_PROPERTY_SIGNATURE = 'ACPS'
APO_CONNECTION_PROPERTY_V2_SIGNATURE = 'ACP2'

# Min and max framerates for the engine
AUDIO_MIN_FRAMERATE = 10.0  # Minimum frame rate for APOs
AUDIO_MAX_FRAMERATE = 384000.0  # Maximum frame rate for APOs

# Min and max # of channels (samples per frame) for the APOs
AUDIO_MIN_CHANNELS = 1  # Current minimum number of channels for APOs
AUDIO_MAX_CHANNELS = 4096  # Current maximum number of channels for APOs


class APO_CONNECTION_BUFFER_TYPE(ENUM):
    APO_CONNECTION_BUFFER_TYPE_ALLOCATED = 0
    APO_CONNECTION_BUFFER_TYPE_EXTERNAL = 1
    APO_CONNECTION_BUFFER_TYPE_DEPENDANT = 2
    APO_CONNECTION_BUFFER_TYPE = 3
    

class APO_CONNECTION_DESCRIPTOR(ctypes.Structure):
    _fields_ = [
        ('Type', APO_CONNECTION_BUFFER_TYPE),
        ('pBuffer', UINT_PTR),
        ('u32MaxFrameCount', UINT32),
        ('pFormat', POINTER(IAudioMediaType)),
        ('u32Signature', UINT32)
    ]


class APO_FLAG(ENUM):
    APO_FLAG_NONE = 0,
    APO_FLAG_INPLACE = 0x1
    APO_FLAG_SAMPLESPERFRAME_MUST_MATCH = 0x2
    APO_FLAG_FRAMESPERSECOND_MUST_MATCH = 0x4
    APO_FLAG_BITSPERSAMPLE_MUST_MATCH = 0x8
    APO_FLAG_MIXER = 0x10
    APO_FLAG_DEFAULT = (
        APO_FLAG_SAMPLESPERFRAME_MUST_MATCH | 
        APO_FLAG_FRAMESPERSECOND_MUST_MATCH | 
        APO_FLAG_BITSPERSAMPLE_MUST_MATCH
    )


class APO_REG_PROPERTIES(ctypes.Structure):
    _fields_ = [
        ('clsid', CLSID),
        ('Flags', APO_FLAG),
        ('szFriendlyName', WCHAR * 256),
        ('szCopyrightInfo', WCHAR * 256),
        ('u32MajorVersion', UINT32),
        ('u32MinorVersion', UINT32),
        ('u32MinInputConnections', UINT32),
        ('u32MaxInputConnections', UINT32),
        ('u32MinOutputConnections', UINT32),
        ('u32MaxOutputConnections', UINT32),
        ('u32MaxInstances', UINT32),
        ('u32NumAPOInterfaces', UINT32),
        ('iidAPOInterfaceList', IID * 1)
    ]


PAPO_REG_PROPERTIES = POINTER(APO_REG_PROPERTIES)


class APOInitBaseStruct(ctypes.Structure):
    _fields_ = [
        ('cbSize', UINT32),
        ('clsid', CLSID)
    ]


class AUDIO_FLOW_TYPE(ENUM):
    AUDIO_FLOW_PULL = 0
    AUDIO_FLOW_PUSH = AUDIO_FLOW_PULL + 1


class EAudioConstriction(ENUM):
    eAudioConstrictionOff = 0
    eAudioConstriction48_16 = eAudioConstrictionOff + 1
    eAudioConstriction44_16 = eAudioConstriction48_16 + 1
    eAudioConstriction14_14 = eAudioConstriction44_16 + 1
    eAudioConstrictionMute = eAudioConstriction14_14 + 1


IID_IAudioProcessingObjectRT = IID(
    '{9E1D6A6D-DDBC-4E95-A4C7-AD64BA37846C}'
)


class IAudioProcessingObjectRT(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioProcessingObjectRT
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'APOProcess',
            (['in'], UINT32, 'stgmAccess'),
            (
                ['in'],
                POINTER(POINTER(APO_CONNECTION_PROPERTY)),
                'ppInputConnections'
            ),
            (['in'], UINT32, 'u32NumOutputConnections'),
            (
                ['in', 'out'],
                POINTER(POINTER(APO_CONNECTION_PROPERTY)),
                'ppOutputConnections'
            ),
        ),
        COMMETHOD(
            [],
            HRESULT,
            'CalcInputFrames',
            (['in'], UINT32, 'u32OutputFrameCount')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'CalcOutputFrames',
            (['in'], UINT32, 'u32InputFrameCount')
        ),
    )


IID_IAudioProcessingObjectVBR = IID(
    '{7ba1db8f-78ad-49cd-9591-f79d80a17c81}'
)


class IAudioProcessingObjectVBR(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioProcessingObjectVBR
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'CalcMaxInputFrames',
            (['in'], UINT32, 'u32MaxOutputFrameCount'),
            (['out'], POINTER(UINT32), 'pu32InputFrameCount')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'CalcMaxOutputFrames',
            (['in'], UINT32, 'u32MaxInputFrameCount'),
            (['out'], POINTER(UINT32), 'pu32OutputFrameCount')
        ),
    )


IID_IAudioProcessingObjectConfiguration = IID(
    '{0E5ED805-ABA6-49c3-8F9A-2B8C889C4FA8}'
)


class IAudioProcessingObjectConfiguration(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioProcessingObjectConfiguration
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'LockForProcess',
            (['in'], UINT32, 'u32NumInputConnections'),
            (
                ['in'],
                POINTER(POINTER(APO_CONNECTION_DESCRIPTOR)),
                'ppInputConnections'
            ),
            (['in'], UINT32, 'u32NumOutputConnections'),
            (
                ['out'],
                POINTER(POINTER(APO_CONNECTION_DESCRIPTOR)),
                'ppOutputConnections'
            )
        ),
        COMMETHOD(
            [],
            HRESULT,
            'UnlockForProcess'
        )
    )


IID_IAudioProcessingObject = IID(
    '{FD7F2B29-24D0-4b5c-B177-592C39F9CA10}'
)


class IAudioProcessingObject(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioProcessingObject
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'Reset'
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetLatency',
            (['out'], POINTER(HNSTIME), 'pTime')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetRegistrationProperties',
            (['out'], POINTER(POINTER(APO_REG_PROPERTIES)), 'ppRegProps')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'Initialize',
            (['in'], UINT32, 'cbDataSize'),
            (['in'], POINTER(BYTE), 'pbyData')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'IsInputFormatSupported',
            (['in'], POINTER(IAudioMediaType), 'pOppositeFormat'),
            (['in'], POINTER(IAudioMediaType), 'pRequestedInputFormat'),
            (
                ['out'],
                POINTER(POINTER(IAudioMediaType)),
                'ppSupportedInputFormat'
            )
        ),
        COMMETHOD(
            [],
            HRESULT,
            'IsOutputFormatSupported',
            (['in'], POINTER(IAudioMediaType), 'pOppositeFormat'),
            (['in'], POINTER(IAudioMediaType), 'pRequestedOutputFormat'),
            (
                ['out'],
                POINTER(POINTER(IAudioMediaType)),
                'ppSupportedOutputFormat'
            )
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetInputChannelCount',
            (['out'], POINTER(UINT32), 'pu32ChannelCount')
        ),
    )


IID_IAudioDeviceModulesClient = IID(
    '{98F37DAC-D0B6-49F5-896A-AA4D169A4C48}'
)


class IAudioDeviceModulesClient(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioDeviceModulesClient
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'SetAudioDeviceModulesManager',
            (['in'], POINTER(comtypes.IUnknown), 'pAudioDeviceModulesManager')
        ),
    )


# APO registration functions
#
# typedef HRESULT (WINAPI FNAPONOTIFICATIONCALLBACK)
# (APO_REG_PROPERTIES* pProperties, VOID* pvRefData);
# extern HRESULT WINAPI RegisterAPO(APO_REG_PROPERTIES const* pProperties);
# extern HRESULT WINAPI UnregisterAPO(REFCLSID clsid);
# extern HRESULT WINAPI RegisterAPONotification(HANDLE hEvent);
# extern HRESULT WINAPI UnregisterAPONotification(HANDLE hEvent);
# extern HRESULT WINAPI EnumerateAPOs
# (FNAPONOTIFICATIONCALLBACK pfnCallback, PVOID pvRefData);
# extern HRESULT WINAPI GetAPOProperties
# (REFCLSID clsid, APO_REG_PROPERTIES* pProperties);

IID_IAudioSystemEffects = IID(
    '{5FA00F27-ADD6-499a-8A9D-6B98521FA75B}'
)


class IAudioSystemEffects(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioSystemEffects
    _methods_ = ()


IID_IAudioSystemEffects2 = IID(
    '{BAFE99D2-7436-44CE-9E0E-4D89AFBFFF56}'
)


class IAudioSystemEffects2(IAudioSystemEffects):
    _case_insensitive_ = False
    _iid_ = IID_IAudioSystemEffects2
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetEffectsList',
            (['out'], POINTER(LPGUID), 'ppEffectsIds'),
            (['out'], POINTER(UINT), 'pcEffects'),
            (['in'], HANDLE, 'Event')
        ),
    )


IID_IAudioSystemEffectsCustomFormats = IID(
    '{B1176E34-BB7F-4f05-BEBD-1B18A534E097}'
)


class IAudioSystemEffectsCustomFormats(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioSystemEffectsCustomFormats
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetFormatCount',
            (['out'], POINTER(UINT), 'pcFormats')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetFormat',
            (['in'], UINT, 'nFormat'),
            (['out'], POINTER(POINTER(IAudioMediaType)), 'ppFormat')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetFormatRepresentation',
            (['in'], UINT, 'nFormat'),
            (['out'], POINTER(LPWSTR), 'ppwstrFormatRep')
        )
    )


IID_IApoAuxiliaryInputConfiguration = IID(
    '{4CEB0AAB-FA19-48ED-A857-87771AE1B768}'
)


class IApoAuxiliaryInputConfiguration(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IApoAuxiliaryInputConfiguration
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'AddAuxiliaryInput',
            (['in'], DWORD, 'dwInputId'),
            (['in'], UINT32, 'cbDataSize'),
            (['in'], POINTER(BYTE), 'pbyData'),
            (['in'], POINTER(APO_CONNECTION_DESCRIPTOR), 'pInputConnection')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'RemoveAuxiliaryInput',
            (['in'], DWORD, 'dwInputId'),
        ),
        COMMETHOD(
            [],
            HRESULT,
            'IsInputFormatSupported',
            (['in'], POINTER(IAudioMediaType), 'pRequestedInputFormat'),
            (
                ['out'],
                POINTER(POINTER(IAudioMediaType)),
                'ppSupportedInputFormat'
            )
        ),
    )


IID_IApoAuxiliaryInputRT = IID(
    '{F851809C-C177-49A0-B1B2-B66F017943AB}'
)


class IApoAuxiliaryInputRT(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IApoAuxiliaryInputRT
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'AcceptInput',
            (['in'], DWORD, 'dwInputId'),
            (['in'], POINTER(APO_CONNECTION_PROPERTY), 'pInputConnection')
        ),
    )


IID_IApoAcousticEchoCancellation = IID(
    '{25385759-3236-4101-A943-25693DFB5D2D}'
)


class IApoAcousticEchoCancellation(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IApoAcousticEchoCancellation
    _methods_ = ()


class APOInitSystemEffects(ctypes.Structure):
    _fields_ = [
        ('APOInit;', APOInitBaseStruct),
        ('pAPOEndpointProperties;', POINTER(IPropertyStore)),
        ('pAPOSystemEffectsProperties;', POINTER(IPropertyStore)),
        ('pReserved;', POINTER(VOID)),
        ('pDeviceCollection;', POINTER(IMMDeviceCollection))
    ]


class APOInitSystemEffects2(ctypes.Structure):
    _fields_ = [
        ('APOInit', APOInitBaseStruct),
        ('pAPOEndpointProperties', POINTER(IPropertyStore)),
        ('pAPOSystemEffectsProperties', POINTER(IPropertyStore)),
        ('pReserved', POINTER(VOID)),
        ('pDeviceCollection', POINTER(IMMDeviceCollection)),
        ('nSoftwareIoDeviceInCollection', UINT),
        ('nSoftwareIoConnectorIndex', UINT),
        ('AudioProcessingMode', GUID),
        ('InitializeForDiscoveryOnly', BOOL)
    ]


class __MIDL___MIDL_itf_audioenginebaseapo_0000_0011_0001(ctypes.Structure):
    _fields_ = [
        ('AddPageParam', LPARAM),
        ('pwstrEndpointID', LPWSTR),
        ('pFxProperties', POINTER(IPropertyStore))
    ]


AudioFXExtensionParams = __MIDL___MIDL_itf_audioenginebaseapo_0000_0011_0001


class NodeTypePropertyKey(PK):

    def __init__(self, l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8, key):
        guid = GUID(l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8)
        super().__init__(guid, key)

    def get(self, endpoint):
        data = self.decode(endpoint, None)
        return data

    def decode(self, endpoint, data):
        return KSNODETYPE.get(data, data)


PKEY_AudioEndpoint_Association = NodeTypePropertyKey(
    0x1da5d803,
    0xd492,
    0x4edd,
    0x8c,
    0x23,
    0xe0,
    0xc0,
    0xff,
    0xee,
    0x7f,
    0x0e,
    2
)


PKEY_FX_Association = NodeTypePropertyKey(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    0
)

PKEY_FX_PreMixEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    1
)

PKEY_FX_PostMixEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    2
)

PKEY_FX_UserInterfaceClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    3
)

PKEY_FX_FriendlyName = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    4
)

PKEY_FX_StreamEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    5
)

PKEY_FX_ModeEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    6
)

PKEY_FX_EndpointEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    7
)

PKEY_FX_KeywordDetector_StreamEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    8
)

PKEY_FX_KeywordDetector_ModeEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    9
)

PKEY_FX_KeywordDetector_EndpointEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    10
)

PKEY_FX_Offload_StreamEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    11
)

PKEY_FX_Offload_ModeEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    12
)

PKEY_CompositeFX_StreamEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    13
)

PKEY_CompositeFX_ModeEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    14
)

PKEY_CompositeFX_EndpointEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    15
)

PKEY_CompositeFX_KeywordDetector_StreamEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    16
)


PKEY_CompositeFX_KeywordDetector_ModeEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    17
)

PKEY_CompositeFX_KeywordDetector_EndpointEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    18
)

PKEY_CompositeFX_Offload_StreamEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    19
)

PKEY_CompositeFX_Offload_ModeEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4fb6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    20
)

PKEY_SFX_ProcessingModes_Supported_For_Streaming = DEFINE_PROPERTYKEY(
    0xd3993a3f,
    0x99c2,
    0x4402,
    0xb5,
    0xec,
    0xa9,
    0x2a,
    0x3,
    0x67,
    0x66,
    0x4b,
    5
)

PKEY_MFX_ProcessingModes_Supported_For_Streaming = DEFINE_PROPERTYKEY(
    0xd3993a3f,
    0x99c2,
    0x4402,
    0xb5,
    0xec,
    0xa9,
    0x2a,
    0x3,
    0x67,
    0x66,
    0x4b,
    6
)

PKEY_EFX_ProcessingModes_Supported_For_Streaming = DEFINE_PROPERTYKEY(
    0xd3993a3f,
    0x99c2,
    0x4402,
    0xb5,
    0xec,
    0xa9,
    0x2a,
    0x3,
    0x67,
    0x66,
    0x4b,
    7
)

PKEY_SFX_KeywordDetector_ProcessingModes_Supported_For_Streaming = (
    DEFINE_PROPERTYKEY(
        0xd3993a3f,
        0x99c2,
        0x4402,
        0xb5,
        0xec,
        0xa9,
        0x2a,
        0x3,
        0x67,
        0x66,
        0x4b,
        8
    )
)
PKEY_MFX_KeywordDetector_ProcessingModes_Supported_For_Streaming = (
    DEFINE_PROPERTYKEY(
        0xd3993a3f,
        0x99c2,
        0x4402,
        0xb5,
        0xec,
        0xa9,
        0x2a,
        0x3,
        0x67,
        0x66,
        0x4b,
        9
    )
)

PKEY_EFX_KeywordDetector_ProcessingModes_Supported_For_Streaming = (
    DEFINE_PROPERTYKEY(
        0xd3993a3f,
        0x99c2,
        0x4402,
        0xb5,
        0xec,
        0xa9,
        0x2a,
        0x3,
        0x67,
        0x66,
        0x4b,
        10
    )
)
PKEY_SFX_Offload_ProcessingModes_Supported_For_Streaming = DEFINE_PROPERTYKEY(
    0xd3993a3f,
    0x99c2,
    0x4402,
    0xb5,
    0xec,
    0xa9,
    0x2a,
    0x3,
    0x67,
    0x66,
    0x4b,
    11
)
PKEY_MFX_Offload_ProcessingModes_Supported_For_Streaming = DEFINE_PROPERTYKEY(
    0xd3993a3f,
    0x99c2,
    0x4402,
    0xb5,
    0xec,
    0xa9,
    0x2a,
    0x3,
    0x67,
    0x66,
    0x4b,
    12
)

PKEY_APO_SWFallback_ProcessingModes = DEFINE_PROPERTYKEY(
    0xd3993a3f,
    0x99c2,
    0x4402,
    0xb5,
    0xec,
    0xa9,
    0x2a,
    0x3,
    0x67,
    0x66,
    0x4b,
    13
)


PKEY_AudioEngine_OEMPeriod = "{E4870E26-3CC5-4CD2-BA46-CA0A9A70ED04},6"


# PKEY_SFX_ProcessingModes_Supported_For_Streaming
# PKEY_MFX_ProcessingModes_Supported_For_Streaming
# PKEY_EFX_ProcessingModes_Supported_For_Streaming
# PKEY_SFX_KeywordDetector_ProcessingModes_Supported_For_Streaming
# PKEY_MFX_KeywordDetector_ProcessingModes_Supported_For_Streaming
# PKEY_EFX_KeywordDetector_ProcessingModes_Supported_For_Streaming
AUDIO_SIGNALPROCESSINGMODE_DEFAULT = GUID("{C18E2F7E-933D-4965-B7D1-1EEF228D2AF3}")
AUDIO_SIGNALPROCESSINGMODE_SPEECH = GUID("{FC1CFC9B-B9D6-4CFA-B5E0-4BB2166878B2}")
AUDIO_SIGNALPROCESSINGMODE_COMMUNICATION = GUID("{98951333-B9CD-48B1-A0A3-FF40682D73F7}")
AUDIO_SIGNALPROCESSINGMODE_MEDIA = GUID("{4780004E-7133-41D8-8C74-660DADD2C0EE}")
AUDIO_SIGNALPROCESSINGMODE_MOVIE = GUID("{B26FEB0D-EC94-477C-9494-D1AB8E753F6E}")
AUDIO_SIGNALPROCESSINGMODE_NOTIFICATION = GUID("{9CF2A70B-F377-403B-BD6B-360863E0355C}")


# PKEY_FX_StreamEffectClsid
# PKEY_FX_ModeEffectClsid
# PKEY_FX_EndpointEffectClsid
# PKEY_FX_KeywordDetector_StreamEffectClsid
# PKEY_FX_KeywordDetector_ModeEffectClsid
# PKEY_FX_KeywordDetector_EndpointEffectClsid
# PKEY_FX_Offload_StreamEffectClsid
# PKEY_FX_Offload_ModeEffectClsid
# PKEY_CompositeFX_StreamEffectClsid
# PKEY_CompositeFX_ModeEffectClsid
# PKEY_CompositeFX_EndpointEffectClsid
# PKEY_CompositeFX_KeywordDetector_StreamEffectClsid
# PKEY_CompositeFX_KeywordDetector_ModeEffectClsid
# PKEY_CompositeFX_KeywordDetector_EndpointEffectClsid
# PKEY_CompositeFX_Offload_StreamEffectClsid
# PKEY_CompositeFX_Offload_ModeEffectClsid
CMpeg4DecMediaObject = GUID('{F371728A-6052-4D47-827C-D039335DFE0A}')
CMpeg43DecMediaObject = GUID('{CBA9E78B-49A3-49EA-93D4-6BCBA8C4DE07}')
CMpeg4sDecMediaObject = GUID('{2A11BAE2-FE6E-4249-864B-9E9ED6E8DBC2}')
CMpeg4sDecMFT = GUID('{5686A0D9-FE39-409F-9DFF-3FDBC849F9F5}')
CZuneM4S2DecMediaObject = GUID('{C56FC25C-0FC6-404A-9503-B10BF51A8AB9}')
CMpeg4EncMediaObject = GUID('{24F258D8-C651-4042-93E4-CA654ABB682C}')
CMpeg4sEncMediaObject = GUID('{6EC5A7BE-D81E-4F9E-ADA3-CD1BF262B6D8}')
CMSSCDecMediaObject = GUID('{7BAFB3B1-D8F4-4279-9253-27DA423108DE}')
CMSSCEncMediaObject = GUID('{8CB9CC06-D139-4AE6-8BB4-41E612E141D5}')
CMSSCEncMediaObject2 = GUID('{F7FFE0A0-A4F5-44B5-949E-15ED2BC66F9D}')
CWMADecMediaObject = GUID('{2EEB4ADF-4578-4D10-BCA7-BB955F56320A}')
CWMAEncMediaObject = GUID('{70F598E9-F4AB-495A-99E2-A7C4D3D89ABF}')
CWMATransMediaObject = GUID('{EDCAD9CB-3127-40DF-B527-0152CCB3F6F5}')
CWMSPDecMediaObject = GUID('{874131CB-4ECC-443B-8948-746B89595D20}')
CWMSPEncMediaObject = GUID('{67841B03-C689-4188-AD3F-4C9EBEEC710B}')
CWMSPEncMediaObject2 = GUID('{1F1F4E1A-2252-4063-84BB-EEE75F8856D5}')
CWMTDecMediaObject = GUID('{F9DBC64E-2DD0-45DD-9B52-66642EF94431}')
CWMTEncMediaObject = GUID('{60B67652-E46B-4E44-8609-F74BFFDC083C}')
CWMVDecMediaObject = GUID('{82D353DF-90BD-4382-8BC2-3F6192B76E34}')
CWMVEncMediaObject2 = GUID('{96B57CDD-8966-410C-BB1F-C97EEA765C04}')
CWMVXEncMediaObject = GUID('{7E320092-596A-41B2-BBEB-175D10504EB6}')
CWMV9EncMediaObject = GUID('{D23B90D0-144F-46BD-841D-59E4EB19DC59}')
CWVC1DecMediaObject = GUID('{C9BFBCCF-E60E-4588-A3DF-5A03B1FD9585}')
CWVC1EncMediaObject = GUID('{44653D0D-8CCA-41E7-BACA-884337B747AC}')
CDeColorConvMediaObject = GUID('{49034C05-F43C-400F-84C1-90A683195A3A}')
CDVDecoderMediaObject = GUID('{E54709C5-1E17-4C8D-94E7-478940433584}')
CDVEncoderMediaObject = GUID('{C82AE729-C327-4CCE-914D-8171FEFEBEFB}')
CMpeg2DecMediaObject = GUID('{863D66CD-CDCE-4617-B47F-C8929CFC28A6}')
CPK_DS_MPEG2Decoder = GUID('{9910C5CD-95C9-4E06-865A-EFA1C8016BF4}')
CAC3DecMediaObject = GUID('{03D7C802-ECFA-47D9-B268-5FB3E310DEE4}')
CPK_DS_AC3Decoder = GUID('{6C9C69D6-0FFC-4481-AFDB-CDF1C79C6F3E}')
CMP3DecMediaObject = GUID('{BBEEA841-0A63-4F52-A7AB-A9B3A84ED38A}')
CResamplerMediaObject = GUID('{F447B69E-1884-4A7E-8055-346F74D6EDB3}')
CResizerMediaObject = GUID('{D3EC8B8B-7728-4FD8-9FE0-7B67D19F73A3}')
CInterlaceMediaObject = GUID('{B5A89C80-4901-407B-9ABC-90D9A644BB46}')
CWMAudioLFXAPO = GUID('{62DC1A93-AE24-464C-A43E-452F824C4250}')
CWMAudioGFXAPO = GUID('{637C490D-EEE3-4C0A-973F-371958802DA2}')
CWMAudioSpdTxDMO = GUID('{5210F8E4-B0BB-47C3-A8D9-7B2282CC79ED}')
CWMAudioAEC = GUID('{745057C7-F353-4F2D-A7EE-58434477730E}')
CClusterDetectorDmo = GUID('{36E820C4-165A-4521-863C-619E1160D4D4}')
CColorControlDmo = GUID('{798059F0-89CA-4160-B325-AEB48EFE4F9A}')
CColorConvertDMO = GUID('{98230571-0087-4204-B020-3282538E57D3}')
CColorLegalizerDmo = GUID('{FDFAA753-E48E-4E33-9C74-98A27FC6726A}')
CFrameInterpDMO = GUID('{0A7CFE1B-6AB5-4334-9ED8-3F97CB37DAA1}')
CFrameRateConvertDmo = GUID('{01F36CE2-0907-4D8B-979D-F151BE91C883}')
CResizerDMO = GUID('{1EA1EA14-48F4-4054-AD1A-E8AEE10AC805}')
CShotDetectorDmo = GUID('{56AEFACD-110C-4397-9292-B0A0C61B6750}')
CSmpteTransformsDmo = GUID('{BDE6388B-DA25-485D-BA7F-FABC28B20318}')
CThumbnailGeneratorDmo = GUID('{559C6BAD-1EA8-4963-A087-8A6810F9218B}')
CTocGeneratorDmo = GUID('{4DDA1941-77A0-4FB1-A518-E2185041D70C}')
CMPEGAACDecMediaObject = GUID('{8DDE1772-EDAD-41C3-B4BE-1F30FB4EE0D6}')
CNokiaAACDecMediaObject = GUID('{3CB2BDE4-4E29-4C44-A73E-2D7C2C46D6EC}')
CVodafoneAACDecMediaObject = GUID('{7F36F942-DCF3-4D82-9289-5B1820278F7C}')
CZuneAACCCDecMediaObject = GUID('{A74E98F2-52D6-4B4E-885B-E0A6CA4F187A}')
CNokiaAACCCDecMediaObject = GUID('{EABF7A6F-CCBA-4D60-8620-B152CC977263}')
CVodafoneAACCCDecMediaObject = GUID('{7E76BF7F-C993-4E26-8FAB-470A70C0D59C}')
CMPEG2EncoderDS = GUID('{5F5AFF4A-2F7F-4279-88C2-CD88EB39D144}')
CMPEG2EncoderVideoDS = GUID('{42150CD9-CA9A-4EA5-9939-30EE037F6E74}')
CMPEG2EncoderAudioDS = GUID('{ACD453BC-C58A-44D1-BBF5-BFB325BE2D78}')
CMPEG2AudDecoderDS = GUID('{E1F1A0B8-BEEE-490D-BA7C-066C40B5E2B9}')
CMPEG2VidDecoderDS = GUID('{212690FB-83E5-4526-8FD7-74478B7939CD}')
CDTVAudDecoderDS = GUID('{8E269032-FE03-4753-9B17-18253C21722E}')
CDTVVidDecoderDS = GUID('{64777DC8-4E24-4BEB-9D19-60A35BE1DAAF}')
CMSAC3Enc = GUID('{C6B400E2-20A7-4E58-A2FE-24619682CE6C}')
CMSH264DecoderMFT = GUID('{62CE7E72-4C71-4D20-B15D-452831A87D9D}')
CMSH264EncoderMFT = GUID('{6CA50344-051A-4DED-9779-A43305165E35}')
CMSH264RemuxMFT = GUID('{05A47EBB-8BF0-4CBF-AD2F-3B71D75866F5}')
CMSAACDecMFT = GUID('{32D186A7-218F-4C75-8876-DD77273A8999}')
AACMFTEncoder = GUID('{93AF0C51-2275-45D2-A35B-F2BA21CAED00}')
CMSDDPlusDecMFT = GUID('{177C0AFE-900B-48D4-9E4C-57ADD250B3D4}')
CMPEG2VideoEncoderMFT = GUID('{E6335F02-80B7-4DC4-ADFA-DFE7210D20D5}')
CMPEG2AudioEncoderMFT = GUID('{46A4DD5C-73F8-4304-94DF-308F760974F4}')
CMSMPEGDecoderMFT = GUID('{2D709E52-123F-49B5-9CBC-9AF5CDE28FB9}')
CMSMPEGAudDecMFT = GUID('{70707B39-B2CA-4015-ABEA-F8447D22D88B}')
CMSDolbyDigitalEncMFT = GUID('{AC3315C9-F481-45D7-826C-0b406C1F64B8}')
MP3ACMCodecWrapper = GUID('{11103421-354C-4CCA-A7A3-1AFF9A5B6701}')
CMSVideoDSPMFT = GUID('{51571744-7FE4-4FF2-A498-2DC34FF74F1B}')