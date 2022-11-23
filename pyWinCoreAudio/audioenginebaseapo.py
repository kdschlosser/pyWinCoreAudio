
from .mmdeviceapi import *
from .propsys import IPropertyStore
from .AudioAPOTypes import *
from .audiomediatype import *
from .data_types import *
from .guiddef import PROPERTYKEY as _PROPERTYKEY
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


PKEY_FX_Association = DEFINE_PROPERTYKEY(
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
