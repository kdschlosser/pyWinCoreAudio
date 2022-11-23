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
import ctypes
import comtypes
from comtypes import CoClass
from typing import Union
from . import utils
from .ksproxy import IKsControl
from .endpointvolumeapi import (
    IAudioEndpointVolumeEx,
    IAudioEndpointVolume
)
from .signal import (
    tw,
    ON_DEVICE_STATE_CHANGED,
    ON_DEVICE_ADDED,
    ON_DEVICE_REMOVED,
    ON_ENDPOINT_DEFAULT_CHANGED,
    ON_DEVICE_PROPERTY_CHANGED
)
from .ks import (
    KSPROPERTY_TYPE_GET,
    KSPROPERTY_TYPE_SET,
    KSPROPERTY_TYPE_TOPOLOGY,
)
from .dsound import (
    DS3DALG_DEFAULT,
    DS3DALG_NO_VIRTUALIZATION,
    DS3DALG_HRTF_FULL,
    DS3DALG_HRTF_LIGHT
)
from .propsys import (
    PROPERTYKEY,
    PIPropertyStore
)

from .propidl import PPROPVARIANT

from .constant import (
    S_OK,
    STGM_READ,
    STGM_WRITE,
)
from .functiondiscoverykeys_devpkey import (
    PKEY_DeviceInterface_FriendlyName,
    PKEY_Device_FriendlyName,
    PKEY_Device_DeviceDesc
)
from .ksmedia import (
    KSNODEPROPERTY_AUDIO_CHANNEL,
    KSNODEPROPERTY,
    AudioSpeakers,
    KSPROPERTY_AEC_MODE,
    KSPROPSETID_Acoustic_Echo_Cancel,
    KSNODETYPE_ACOUSTIC_ECHO_CANCEL,
    AEC_MODE_PASS_THROUGH,
    AEC_MODE_HALF_DUPLEX,
    AEC_MODE_FULL_DUPLEX,
    KSPROPSETID_Audio,
    PARTID_MASK,
    KSNODETYPE_TONE,
    KSNODETYPE_EQUALIZER,
    KSNODETYPE_AGC,
    # KSNODETYPE_NOISE_SUPPRESS,
    KSNODETYPE_LOUDNESS,
    # KSNODETYPE_PROLOGIC_DECODER,
    KSNODETYPE_STEREO_WIDE,
    KSNODETYPE_REVERB,
    KSNODETYPE_CHORUS,
    KSNODETYPE_3D_EFFECTS,
    # KSNODETYPE_PARAMETRIC_EQUALIZER,
    # KSNODETYPE_DYN_RANGE_COMPRESSOR,
    # KSPROPERTY_AUDIO_DYNAMIC_RANGE,
    KSPROPERTY_AUDIO_BASS,
    KSPROPERTY_AUDIO_MID,
    KSPROPERTY_AUDIO_TREBLE,
    KSPROPERTY_AUDIO_BASS_BOOST,
    KSPROPERTY_AUDIO_EQ_LEVEL,
    KSPROPERTY_AUDIO_NUM_EQ_BANDS,
    KSPROPERTY_AUDIO_EQ_BANDS,
    KSPROPERTY_AUDIO_AGC,
    KSPROPERTY_AUDIO_LOUDNESS,
    KSPROPERTY_AUDIO_WIDENESS,
    KSPROPERTY_AUDIO_REVERB_LEVEL,
    KSPROPERTY_AUDIO_CHORUS_LEVEL,
    # KSPROPERTY_AUDIO_SURROUND_ENCODE,
    # KSPROPERTY_AUDIO_3D_INTERFACE,
    # KSPROPERTY_AUDIO_PEQ_MAX_BANDS,
    # KSPROPERTY_AUDIO_PEQ_NUM_BANDS,
    # KSPROPERTY_AUDIO_PEQ_BAND_CENTER_FREQ,
    # KSPROPERTY_AUDIO_PEQ_BAND_Q_FACTOR,
    # KSPROPERTY_AUDIO_PEQ_BAND_LEVEL,
    KSPROPERTY_AUDIO_CHORUS_MODULATION_RATE,
    KSPROPERTY_AUDIO_CHORUS_MODULATION_DEPTH,
    KSPROPERTY_AUDIO_REVERB_TIME,
    KSPROPERTY_AUDIO_REVERB_DELAY_FEEDBACK,
    # KSPROPERTY_AUDIO_MIC_SENSITIVITY,
    # KSPROPERTY_AUDIO_MIC_SNR,
    # KSPROPERTY_AUDIO_MIC_SENSITIVITY2,
    JackDescription,
    EPcxConnectionType,
    KSJACK_SINK_INFORMATION,
    KSNODETYPE
)


IID_IMMDeviceCollection = IID(
    '{0BD7A1BE-7A1A-44DB-8397-CC5392387B5E}'
)
IID_IMMDeviceEnumerator = IID(
    '{A95664D2-9614-4F35-A746-DE8DB63617E6}'
)
IID_IMMDevice = IID(
    '{D666063F-1587-4E43-81F1-B948E807363F}'
)
IID_IMMNotificationClient = IID(
    '{7991EEC9-7E89-4D85-8390-6C703CEC60C0}'
)
IID_IMMEndpoint = IID(
    '{1BE09788-6894-4089-8586-9A2A6C265AC5}'
)
IID_IMMDeviceActivator = IID(
    '{3B0D0EA4-D0A9-4B0E-935B-09516746FAC0}'
)
IID_IActivateAudioInterfaceAsyncOperation = IID(
    '{72A22D78-CDE4-431D-B8CC-843A71199B6D}'
)
IID_IActivateAudioInterfaceCompletionHandler = IID(
    '{41D949AB-9862-444A-80F6-C261334DA5EB}'
)
IID_MMDeviceAPILib = (
    '{2FDAAFA3-7523-4F66-9957-9D5E7FE698F6}'
)
CLSID_MMDeviceEnumerator = IID(
    '{BCDE0395-E52F-467C-8E3D-C4579291692E}'
)


class EndpointFormFactor(ENUM):
    RemoteNetworkDevice = 0
    Speakers = 1
    LineLevel = 2
    Headphones = 3
    Microphone = 4
    Headset = 5
    Handset = 6
    UnknownDigitalPassthrough = 7
    SPDIF = ENUM_VALUE(8, 'SPDIF')
    DigitalAudioDisplayDevice = 9
    UnknownFormFactor = ENUM_VALUE(10, 'Unknown')


PEndpointFormFactor = POINTER(EndpointFormFactor)


class ERole(ENUM):
    eConsole = 0
    eMultimedia = 1
    eCommunications = 2


class EDataFlow(ENUM):
    eRender = ENUM_VALUE(0, 'Render')
    eCapture = ENUM_VALUE(1, 'Capture')
    eAll = 2


PEDataFlow = POINTER(EDataFlow)


class DataFlow(ENUM):
    In = 0
    Out = 1


PDataFlow = POINTER(DataFlow)


DEVICE_STATE_ACTIVE = 0x00000001
DEVICE_STATE_DISABLED = 0x00000002
DEVICE_STATE_NOTPRESENT = 0x00000004
DEVICE_STATE_UNPLUGGED = 0x00000008
DEVICE_STATE_MASK_ALL = 0x0000000F

from . import policyconfig  # NOQA

from .audioclient import IAudioClient  # NOQA
from .devicetopologyapi import (
    IDeviceTopology,
    IAudioInputSelector,
    IKsJackDescription,
    IKsJackDescription2,
    IKsJackSinkInformation,
    # IAudioAutoGainControl,
    # IAudioBass,
    IAudioChannelConfig,
    # IAudioLoudness,
    # IAudioMidrange,
    IAudioOutputSelector,
    # IAudioTreble,
    IConnector,
    ISubunit,
    PartType,
    ConnectorType
)  # NOQA


_CoTaskMemFree = ctypes.windll.ole32.CoTaskMemFree


from .guiddef import PROPERTYKEY as _PROPERTYKEY  # NOQA


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


PKEY_AudioEndpoint_FormFactor = DEFINE_PROPERTYKEY(
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
    0
)

PKEY_AudioEndpoint_PhysicalSpeakers = DEFINE_PROPERTYKEY(
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
    3
)

PKEY_AudioEndpoint_GUID = DEFINE_PROPERTYKEY(
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
    4
)

PKEY_AudioEndpoint_Disable_SysFx = DEFINE_PROPERTYKEY(
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
    5
)

ENDPOINT_SYSFX_ENABLED = 0x00000000  # System Effects are enabled.
ENDPOINT_SYSFX_DISABLED = 0x00000001  # System Effects are disabled.

PKEY_AudioEndpoint_FullRangeSpeakers = DEFINE_PROPERTYKEY(
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
    6
)

PKEY_AudioEndpoint_JackSubType = DEFINE_PROPERTYKEY(
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
    8
)

PKEY_AudioEndpoint_ControlPanelPageProvider = DEFINE_PROPERTYKEY(
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
    1
)

PKEY_AudioEndpoint_Association = DEFINE_PROPERTYKEY(
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

PKEY_AudioEndpoint_Supports_EventDriven_Mode = DEFINE_PROPERTYKEY(
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
    7
)

PKEY_AudioEndpoint_Default_VolumeInDb = DEFINE_PROPERTYKEY(
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
    9
)

PKEY_AudioEngine_DeviceFormat = DEFINE_PROPERTYKEY(
    0xf19f064d,
    0x82c,
    0x4e27,
    0xbc,
    0x73,
    0x68,
    0x82,
    0xa1,
    0xbb,
    0x8e,
    0x4c,
    0
)

PKEY_AudioEngine_OEMFormat = DEFINE_PROPERTYKEY(
    0xe4870e26,
    0x3cc5,
    0x4cd2,
    0xba,
    0x46,
    0xca,
    0xa,
    0x9a,
    0x70,
    0xed,
    0x4,
    3
)

PKEY_AudioEndpointLogo_IconEffects = DEFINE_PROPERTYKEY(
    0xf1ab780d,
    0x2010,
    0x4ed3,
    0xa3,
    0xa6,
    0x8b,
    0x87,
    0xf0,
    0xf0,
    0xc4,
    0x76,
    0
)

PKEY_AudioEndpointLogo_IconPath = DEFINE_PROPERTYKEY(
    0xf1ab780d,
    0x2010,
    0x4ed3,
    0xa3,
    0xa6,
    0x8b,
    0x87,
    0xf0,
    0xf0,
    0xc4,
    0x76,
    1
)

PKEY_AudioEndpointSettings_MenuText = DEFINE_PROPERTYKEY(
    0x14242002,
    0x0320,
    0x4de4,
    0x95,
    0x55,
    0xa7,
    0xd8,
    0x2b,
    0x73,
    0xc2,
    0x86,
    0
)

PKEY_AudioEndpointSettings_LaunchContract = DEFINE_PROPERTYKEY(
    0x14242002,
    0x0320,
    0x4de4,
    0x95,
    0x55,
    0xa7,
    0xd8,
    0x2b,
    0x73,
    0xc2,
    0x86,
    1
)

# {D04E05A6-594B-4FB6-A80D-01AF5EED7D1D}, 0
PKEY_SYSFX_Association = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4FB6,
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
# VT: None

# {D04E05A6-594B-4FB6-A80D-01AF5EED7D1D}, 1
PKEY_SYSFX_PreMixClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4FB6,
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
# VT: None

# {D04E05A6-594B-4FB6-A80D-01AF5EED7D1D}, 2
PKEY_SYSFX_PostMixClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4FB6,
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
# VT: None

# {D04E05A6-594B-4FB6-A80D-01AF5EED7D1D}, 3
PKEY_SYSFX_UiClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4FB6,
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
# VT: None

# {D04E05A6-594B-4FB6-A80D-01AF5EED7D1D}, 5
PKEY_SYSFX_StreamEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4FB6,
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
# VT: None

# {D04E05A6-594B-4FB6-A80D-01AF5EED7D1D}, 6
PKEY_SYSFX_ModeEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4FB6,
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
# VT: None

# {D04E05A6-594B-4FB6-A80D-01AF5EED7D1D}, 7
PKEY_SYSFX_EndpointEffectClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4FB6,
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
# VT: None

# {D04E05A6-594B-4FB6-A80D-01AF5EED7D1D}, 196611
PKEY_SYSFX_ChainClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4FB6,
    0xA8,
    0x0D,
    0x01,
    0xAF,
    0x5E,
    0xED,
    0x7D,
    0x1D,
    196611
)
# VT: None

# {259abffc-50a7-47ce-af08-68c9a7d73366}, 12
PKEY_AudioEndpoint_Icon = DEFINE_PROPERTYKEY(
    0x259ABFFC,
    0x50A7,
    0x47CE,
    0xAF,
    0x08,
    0x68,
    0xC9,
    0xA7,
    0xD7,
    0x33,
    0x66,
    12
)
# VT: None

# {F3E80BEF-1723-4FF2-BCC4-7F83DC5E46D4}, 4
PKEY_AudioDevice_EnableEndpointByDefault = DEFINE_PROPERTYKEY(
    0xF3E80BEF,
    0x1723,
    0x4FF2,
    0xBC,
    0xC4,
    0x7F,
    0x83,
    0xDC,
    0x5E,
    0x46,
    0xD4,
    4
)
# VT: None

# {D3993A3F-99C2-4402-B5EC-A92A0367664B}, 1
PKEY_LFX_ProcessingModes_Supported_For_Streaming = DEFINE_PROPERTYKEY(
    0xD3993A3F,
    0x99C2,
    0x4402,
    0xB5,
    0xEC,
    0xA9,
    0x2A,
    0x03,
    0x67,
    0x66,
    0x4B,
    1
)
# VT: None

# {1da5d803-d492-4edd-8c23-e0c0ffee7f0e}, 3
PKEY_AudioEngine_PhysicalSpeaker = DEFINE_PROPERTYKEY(
    0x1DA5D803,
    0xD492,
    0x4EDD,
    0x8C,
    0x23,
    0xE0,
    0xC0,
    0xFF,
    0xEE,
    0x7F,
    0x0E,
    3
)
# VT: None

# {1da5d803-d492-4edd-8c23-e0c0ffee7f0e}, 6
PKEY_AudioEngine_FullRangeSpeaker = DEFINE_PROPERTYKEY(
    0x1DA5D803,
    0xD492,
    0x4EDD,
    0x8C,
    0x23,
    0xE0,
    0xC0,
    0xFF,
    0xEE,
    0x7F,
    0x0E,
    6
)
# VT: None

# {F1056047-B091-4d85-A5C0-B13D4D8BAC57}, 0
PKEY_EFFECTNODEINFO_RENDER = DEFINE_PROPERTYKEY(
    0xF1056047,
    0xB091,
    0x4D85,
    0xA5,
    0xC0,
    0xB1,
    0x3D,
    0x4D,
    0x8B,
    0xAC,
    0x57,
    0
)
# VT: None

# {F1056047-B091-4d85-A5C0-B13D4D8BAC57}, 1
PKEY_EFFECTNODEINFO_CAPTURE = DEFINE_PROPERTYKEY(
    0xF1056047,
    0xB091,
    0x4D85,
    0xA5,
    0xC0,
    0xB1,
    0x3D,
    0x4D,
    0x8B,
    0xAC,
    0x57,
    1
)
# VT: None

# {F5376650-918F-4cf4-91FB-D123CF4E1350}, 1
PKEY_CTGUID = DEFINE_PROPERTYKEY(
    0xF5376650,
    0x918F,
    0x4CF4,
    0x91,
    0xFB,
    0xD1,
    0x23,
    0xCF,
    0x4E,
    0x13,
    0x50,
    1
)
# VT: None

# {0D63112A-5B63-4087-9AE2-5AE185B79E43}, 0
PKEY_CT_LINETOAUX_GUID = DEFINE_PROPERTYKEY(
    0x0D63112A,
    0x5B63,
    0x4087,
    0x9A,
    0xE2,
    0x5A,
    0xE1,
    0x85,
    0xB7,
    0x9E,
    0x43,
    0
)
# VT: None

# {D04E05A6-594B-4FB6-A80D-01AF5EED7D1D}, 3
PKEY_FX_UiClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4FB6,
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
# VT: None

# {B725F130-47EF-101A-A5F1-02608C9EEBAC}, 10
PKEY_ItemNameDisplay = DEFINE_PROPERTYKEY(
    0xB725F130,
    0x47EF,
    0x101A,
    0xA5,
    0xF1,
    0x02,
    0x60,
    0x8C,
    0x9E,
    0xEB,
    0xAC,
    10
)
# VT: None

# {D04E05A6-594B-4FB6-A80D-01AF5EED7D1D}, 1
PKEY_FX_PreMixClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4FB6,
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
# VT: None

# {D04E05A6-594B-4FB6-A80D-01AF5EED7D1D}, 2
PKEY_FX_PostMixClsid = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4FB6,
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
# VT: None

# {F3E80BEF-1723-4FF2-BCC4-7F83DC5E46D4}, 3
PKEY_AudioDevice_NeverSetAsDefaultEndpoint = DEFINE_PROPERTYKEY(
    0xF3E80BEF,
    0x1723,
    0x4FF2,
    0xBC,
    0xC4,
    0x7F,
    0x83,
    0xDC,
    0x5E,
    0x46,
    0xD4,
    3
)
# VT: None

# {D04E05A6-594B-4fb6-A80D-01AF5EED7D1D}, 13
PKEY_Composite_SFX = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4FB6,
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
# VT: None

# {D04E05A6-594B-4fb6-A80D-01AF5EED7D1D}, 14
PKEY_Composite_MFX = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4FB6,
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
# VT: None

# {D04E05A6-594B-4fb6-A80D-01AF5EED7D1D}, 15
PKEY_Composite_EFX = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4FB6,
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
# VT: None

# {D04E05A6-594B-4fb6-A80D-01AF5EED7D1D}, 19
PKEY_Composite_Offload_SFX = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4FB6,
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
# VT: None

# {D04E05A6-594B-4fb6-A80D-01AF5EED7D1D}, 20
PKEY_Composite_Offload_MFX = DEFINE_PROPERTYKEY(
    0xD04E05A6,
    0x594B,
    0x4FB6,
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
# VT: None

# {1DA5D803-D492-4EDD-8C23-E0C0FFEE7F0E}, 1
PKEY_AudioEndpoint_ControlPanelProvider = DEFINE_PROPERTYKEY(
    0x1DA5D803,
    0xD492,
    0x4EDD,
    0x8C,
    0x23,
    0xE0,
    0xC0,
    0xFF,
    0xEE,
    0x7F,
    0x0E,
    1
)
# VT: None

# {1DA5D803-D492-4EDD-8C23-E0C0FFEE7F0E}, 1
PKEY_AudioEndpoint_Ext_UiClsid = DEFINE_PROPERTYKEY(
    0x1DA5D803,
    0xD492,
    0x4EDD,
    0x8C,
    0x23,
    0xE0,
    0xC0,
    0xFF,
    0xEE,
    0x7F,
    0x0E,
    1
)
# VT: None

# {A45C254E-DF1C-4EFD-8020-67D146A850E0}, 2
PKEY_AudioEndpoint_Name = DEFINE_PROPERTYKEY(
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
# VT: None

# {0F8412D3-DC5C-4DB3-B174-DC47A859435C}, 0
PKEY_BYPASS_TP_EFFECTS = DEFINE_PROPERTYKEY(
    0x0F8412D3,
    0xDC5C,
    0x4DB3,
    0xB1,
    0x74,
    0xDC,
    0x47,
    0xA8,
    0x59,
    0x43,
    0x5C,
    0
)
# VT: None


class _IMMNotificationClient(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IMMNotificationClient
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'OnDeviceStateChanged',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], DWORD, 'dwNewState'),
        ),
        COMMETHOD(
            [],
            HRESULT,
            'OnDeviceAdded',
            (['in'], LPCWSTR, 'pwstrDeviceId')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'OnDeviceRemoved',
            (['in'], LPCWSTR, 'pwstrDeviceId')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'OnDefaultDeviceChanged',
            (['in'], EDataFlow, 'flow'),
            (['in'], ERole, 'role'),
            (['in'], LPCWSTR, 'pwstrDefaultDeviceId')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'OnPropertyValueChanged',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], PROPERTYKEY, 'key')
        )
    )


# noinspection PyTypeChecker
PIMMNotificationClient = POINTER(_IMMNotificationClient)


class IMMNotificationClient(comtypes.COMObject):
    _com_interfaces_ = [_IMMNotificationClient]

    def __init__(self, device_enum):
        self.__device_enum = device_enum
        self.__last_default_endpoint = []
        comtypes.COMObject.__init__(self)

    def OnDeviceStateChanged(self, pwstrDeviceId, dwNewState):
        # print('OnDeviceStateChanged')
        pwstrDeviceId = utils.convert_to_string(pwstrDeviceId)

        def _do(dev_id, new_state):

            if new_state == DEVICE_STATE_UNPLUGGED:
                state = 'Unplugged'
            elif new_state == DEVICE_STATE_NOTPRESENT:
                state = 'Not Present'
            elif new_state == DEVICE_STATE_DISABLED:
                state = 'Disabled'
            elif new_state == DEVICE_STATE_ACTIVE:
                state = 'Active'
            else:
                state = 'Unknown'

            for device in self.__device_enum:
                if device.id == dev_id:
                    ON_DEVICE_STATE_CHANGED.signal(
                        device=device,
                        new_state=state
                    )
                    break

                for endpoint in device:
                    if endpoint.id == dev_id:
                        ON_DEVICE_STATE_CHANGED.signal(
                            device=device,
                            endpoint=endpoint,
                            new_state=state
                        )
                        break
                else:
                    continue

                break

        tw.add(_do, pwstrDeviceId, dwNewState)

        return S_OK

    def OnDeviceAdded(self, pwstrDeviceId):
        # print('OnDeviceAdded')

        pwstrDeviceId = utils.convert_to_string(pwstrDeviceId)

        def _do(dev_id):
            for device in self.__device_enum:
                if device.id == dev_id:
                    ON_DEVICE_ADDED.signal(device=device)
                    break

        tw.add(_do, pwstrDeviceId)

        return S_OK

    def OnDeviceRemoved(self, pwstrDeviceId):
        # print('OnDeviceRemoved')

        pwstrDeviceId = utils.convert_to_string(pwstrDeviceId)

        def _do(dev_id):
            name = self.__device_enum.get_device_name(dev_id)

            for _ in self.__device_enum:
                continue

            if name:
                ON_DEVICE_REMOVED.signal(name=name)

        tw.add(_do, pwstrDeviceId)
        return S_OK

    def OnDefaultDeviceChanged(self, flow, role, pwstrDefaultDeviceId):
        pwstrDefaultDeviceId = utils.convert_to_string(pwstrDefaultDeviceId)

        flow = EDataFlow.get(flow.value)
        role = ERole.get(role.value)

        def _do(dev_id, flw, rle):
            for device in self.__device_enum:
                for endpoint in device:
                    if endpoint.id == dev_id:
                        break
                else:
                    continue

                break
            else:
                return

            ON_ENDPOINT_DEFAULT_CHANGED.signal(
                device=device,
                endpoint=endpoint,
                role=rle,
                flow=flw
            )

        if (
            (pwstrDefaultDeviceId, flow, role)
            not in self.__last_default_endpoint
        ):
            self.__last_default_endpoint.append(
                (pwstrDefaultDeviceId, flow, role)
            )
            while len(self.__last_default_endpoint) > 3:
                self.__last_default_endpoint.pop(0)

            tw.add(_do, pwstrDefaultDeviceId, flow, role)

        return S_OK

    def OnPropertyValueChanged(self, pwstrDeviceId, key):
        # print('OnPropertyValueChanged')

        pwstrDeviceId = utils.convert_to_string(pwstrDeviceId)

        def _do(dev_id, k):

            for device in self.__device_enum:
                if device.id == dev_id:
                    ON_DEVICE_PROPERTY_CHANGED.signal(device=device, key=k)
                    break

                for endpoint in device:
                    if endpoint.id == dev_id:
                        ON_DEVICE_PROPERTY_CHANGED.signal(
                            device=device,
                            endpoint=endpoint,
                            key=k
                        )
                        break
                else:
                    continue

                break

        tw.add(_do, pwstrDeviceId, key)

        return S_OK


class IActivateAudioInterfaceAsyncOperation(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IActivateAudioInterfaceAsyncOperation
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetActivateResult',
            (['out'], LPHRESULT, 'activateResult'),
            (['out'], POINTER(PIUnknown), 'activatedInterface'),
        ),
    )


# noinspection PyTypeChecker
PIActivateAudioInterfaceAsyncOperation = POINTER(
    IActivateAudioInterfaceAsyncOperation
)


class IActivateAudioInterfaceCompletionHandler(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IActivateAudioInterfaceCompletionHandler
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'ActivateCompleted',
            (
                ['in'],
                PIActivateAudioInterfaceAsyncOperation,
                'activateOperation'
            ),
        ),
    )


# noinspection PyTypeChecker
PIActivateAudioInterfaceCompletionHandler = POINTER(
    IActivateAudioInterfaceCompletionHandler
)


class IMMEndpoint(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IMMEndpoint
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetDataFlow',
            (['out'], PEDataFlow, 'pDataFlow')
        ),
    )


# noinspection PyTypeChecker
PIMMEndpoint = POINTER(IMMEndpoint)
#
# KSNODETYPE_PROLOGIC_DECODER
# KSPROPERTY_AUDIO_SURROUND_ENCODE,
# KSPROPERTY_AUDIO_3D_INTERFACE,
#
#
# KSNODETYPE_DYN_RANGE_COMPRESSOR
# KSPROPERTY_AUDIO_DYNAMIC_RANGE
#
#
# KSNODETYPE_NOISE_SUPPRESS
# KSPROPERTY_AUDIO_MIC_SENSITIVITY
# KSPROPERTY_AUDIO_MIC_SNR
# KSPROPERTY_AUDIO_MIC_SENSITIVITY2
#
#
# KSNODETYPE_PARAMETRIC_EQUALIZER
# KSPROPERTY_AUDIO_PEQ_MAX_BANDS
# KSPROPERTY_AUDIO_PEQ_NUM_BANDS
# KSPROPERTY_AUDIO_PEQ_BAND_CENTER_FREQ
# KSPROPERTY_AUDIO_PEQ_BAND_Q_FACTOR
# KSPROPERTY_AUDIO_PEQ_BAND_LEVEL
#
#


class ChorusReverbBase(object):
    _func1_id = None
    _func2_id = None
    _level_id = None
    _node_type = None

    def __init__(self, device, device_topology, device_enum):
        self.__device = device
        self.__device_topology = device_topology
        self.__device_enum = device_enum

    @property
    def _func1(self):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return False

        ksprop = KSNODEPROPERTY()
        ksprop.Property.Set = KSPROPSETID_Audio
        ksprop.Property.Id = self._func1_id
        ksprop.Property.Flags = KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY

        for subunit in self.__device.subunits:
            part = subunit.part
            if part.sub_type != self._node_type:
                continue

            local_id = part.local_id
            ksprop.NodeId = local_id & PARTID_MASK
            bValue = ULONG()
            valueSize = ULONG()
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            if valueSize.value == 0:
                continue

            return bool(bValue)

        return False

    @_func1.setter
    def _func1(self, value):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return

        ksprop = KSNODEPROPERTY()
        ksprop.Property.Set = KSPROPSETID_Audio
        ksprop.Property.Id = self._func1_id
        ksprop.Property.Flags = KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY

        for subunit in self.__device.subunits:
            part = subunit.part

            if part.sub_type != self._node_type:
                continue

            local_id = part.local_id
            ksprop.NodeId = local_id & PARTID_MASK
            bValue = ULONG(value)
            valueSize = ULONG(ctypes.sizeof(bValue))
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            break

    @property
    def _func2(self):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return False

        ksprop = KSNODEPROPERTY()
        ksprop.Property.Set = KSPROPSETID_Audio
        ksprop.Property.Id = self._func2_id
        ksprop.Property.Flags = KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY

        for subunit in self.__device.subunits:
            part = subunit.part
            if part.sub_type != self._node_type:
                continue

            local_id = part.local_id
            ksprop.NodeId = local_id & PARTID_MASK
            bValue = ULONG()
            valueSize = ULONG()
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            if valueSize.value == 0:
                continue

            return bool(bValue)

        return False

    @_func2.setter
    def _func2(self, value):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return

        ksprop = KSNODEPROPERTY()
        ksprop.Property.Set = KSPROPSETID_Audio
        ksprop.Property.Id = self._func2_id
        ksprop.Property.Flags = KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY

        for subunit in self.__device.subunits:
            part = subunit.part

            if part.sub_type != self._node_type:
                continue

            local_id = part.local_id
            ksprop.NodeId = local_id & PARTID_MASK
            bValue = ULONG(value)
            valueSize = ULONG(ctypes.sizeof(bValue))
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            break

    @property
    def level(self):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return False

        ksprop = KSNODEPROPERTY()
        ksprop.Property.Set = KSPROPSETID_Audio
        ksprop.Property.Id = self._level_id
        ksprop.Property.Flags = KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY

        for subunit in self.__device.subunits:
            part = subunit.part
            if part.sub_type != self._node_type:
                continue

            local_id = part.local_id
            ksprop.NodeId = local_id & PARTID_MASK
            bValue = ULONG()
            valueSize = ULONG()
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            if valueSize.value == 0:
                continue

            return bool(bValue)

        return False

    @level.setter
    def level(self, value):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return

        ksprop = KSNODEPROPERTY()
        ksprop.Property.Set = KSPROPSETID_Audio
        ksprop.Property.Id = self._level_id
        ksprop.Property.Flags = KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY

        for subunit in self.__device.subunits:
            part = subunit.part

            if part.sub_type != self._node_type:
                continue

            local_id = part.local_id
            ksprop.NodeId = local_id & PARTID_MASK
            bValue = ULONG(value)
            valueSize = ULONG(ctypes.sizeof(bValue))
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            break


class Chorus(ChorusReverbBase):
    _func1_id = KSPROPERTY_AUDIO_CHORUS_MODULATION_RATE
    _func2_id = KSPROPERTY_AUDIO_CHORUS_MODULATION_DEPTH
    _level_id = KSPROPERTY_AUDIO_CHORUS_LEVEL
    _node_type = KSNODETYPE_CHORUS

    @property
    def modulation_rate(self):
        return self._func1

    @modulation_rate.setter
    def modulation_rate(self, value):
        self._func1 = value

    @property
    def modulation_depth(self):
        return self._func2

    @modulation_depth.setter
    def modulation_depth(self, value):
        self._func2 = value


class Reverb(ChorusReverbBase):
    _func1_id = KSPROPERTY_AUDIO_REVERB_TIME
    _func2_id = KSPROPERTY_AUDIO_REVERB_DELAY_FEEDBACK
    _level_id = KSPROPERTY_AUDIO_REVERB_LEVEL
    _node_type = KSNODETYPE_REVERB

    @property
    def reverb_time(self):
        return self._func1

    @reverb_time.setter
    def reverb_time(self, value):
        self._func1 = value

    @property
    def delay_feedback(self):
        return self._func2

    @delay_feedback.setter
    def delay_feedback(self, value):
        self._func2 = value


class EQBand(object):

    def __init__(
        self,
        band,
        frequency,
        device,
        channel_num,
        device_topology,
        device_enum
    ):
        self.__band = band
        self.__frequency = frequency
        self.__device = device
        self.__channel_num = channel_num
        self.__device_topology = device_topology
        self.__device_enum = device_enum

    @property
    def band(self):
        return self.__band

    @property
    def frequency(self):
        return self.__frequency

    @property
    def level(self):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if iks_control:
            ksprop = KSNODEPROPERTY_AUDIO_CHANNEL()
            ksprop.NodeProperty.Property.Set = KSPROPSETID_Audio
            ksprop.NodeProperty.Property.Id = KSPROPERTY_AUDIO_NUM_EQ_BANDS
            ksprop.NodeProperty.Property.Flags = (
                KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY
            )
            ksprop.Channel = self.__channel_num

            for subunit in self.__device.subunits:
                part = subunit.part
                if part.sub_type != KSNODETYPE_EQUALIZER:
                    continue

                local_id = part.local_id
                ksprop.NodeProperty.NodeId = local_id & PARTID_MASK
                bValue = ULONG()
                valueSize = ULONG()
                try:
                    iks_control.KsProperty(
                        ctypes.byref(ksprop),
                        ctypes.sizeof(ksprop),
                        ctypes.byref(bValue),
                        ctypes.sizeof(bValue),
                        ctypes.byref(valueSize)
                    )
                except comtypes.COMError:
                    continue

                if valueSize.value == 0:
                    continue

                num_bands = bValue.value
                bValue = (ULONG * num_bands)()
                valueSize = ULONG()
                ksprop.NodeProperty.Property.Id = KSPROPERTY_AUDIO_EQ_LEVEL

                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    bValue,
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )

                return bValue[self.__band].value

    @level.setter
    def level(self, value):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return

        ksprop = KSNODEPROPERTY_AUDIO_CHANNEL()
        ksprop.NodeProperty.Property.Set = KSPROPSETID_Audio
        ksprop.NodeProperty.Property.Id = KSPROPERTY_AUDIO_NUM_EQ_BANDS
        ksprop.NodeProperty.Property.Flags = (
           KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY
        )
        ksprop.Channel = self.__channel_num

        for subunit in self.__device.subunits:
            part = subunit.part
            if part.sub_type != KSNODETYPE_EQUALIZER:
                continue

            local_id = part.local_id
            ksprop.NodeProperty.NodeId = local_id & PARTID_MASK
            bValue = ULONG()
            valueSize = ULONG()
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            if valueSize.value == 0:
                continue

            num_bands = bValue.value
            bValue = (ULONG * num_bands)()
            valueSize = ULONG()
            ksprop.NodeProperty.Property.Id = KSPROPERTY_AUDIO_EQ_LEVEL

            iks_control.KsProperty(
                ctypes.byref(ksprop),
                ctypes.sizeof(ksprop),
                bValue,
                ctypes.sizeof(bValue),
                ctypes.byref(valueSize)
            )

            bValue[self.__band] = value

            ksprop.NodeProperty.Property.Flags = (
                KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY
            )
            valueSize = ULONG(ctypes.sizeof(bValue))

            iks_control.KsProperty(
                ctypes.byref(ksprop),
                ctypes.sizeof(ksprop),
                bValue,
                ctypes.sizeof(bValue),
                ctypes.byref(valueSize)
            )


class Equalizer(object):
    def __init__(self, channel_num, device, device_topology, device_enum):
        self.__device = device
        self.__channel_num = channel_num
        self.__device_topology = device_topology
        self.__device_enum = device_enum

    @property
    def bands(self):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)

        res = []

        if not iks_control:
            return res

        ksprop = KSNODEPROPERTY_AUDIO_CHANNEL()
        ksprop.NodeProperty.Property.Set = KSPROPSETID_Audio
        ksprop.NodeProperty.Property.Id = KSPROPERTY_AUDIO_NUM_EQ_BANDS
        ksprop.NodeProperty.Property.Flags = (
            KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY
        )
        ksprop.Channel = self.__channel_num

        for subunit in self.__device.subunits:
            part = subunit.part
            if part.sub_type != KSNODETYPE_EQUALIZER:
                continue

            local_id = part.local_id
            ksprop.NodeProperty.NodeId = local_id & PARTID_MASK
            bValue = ULONG()
            valueSize = ULONG()
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            if valueSize.value == 0:
                continue

            num_bands = bValue.value
            bValue = (ULONG * num_bands)()
            valueSize = ULONG()
            ksprop.NodeProperty.Property.Id = KSPROPERTY_AUDIO_EQ_BANDS

            iks_control.KsProperty(
                ctypes.byref(ksprop),
                ctypes.sizeof(ksprop),
                bValue,
                ctypes.sizeof(bValue),
                ctypes.byref(valueSize)
            )

            for i in range(num_bands):
                frequency = bValue[i].value
                eq_band = EQBand(
                    i,
                    frequency,
                    self.__device,
                    self.__channel_num,
                    self.__device_topology,
                    self.__device_enum
                )

                res.append(eq_band)

            break

        return res


class AudioChannel(object):

    def __init__(self, device, channel_num, device_topology, device_enum):
        self.__device = device
        self.__channel_num = channel_num
        self.__device_topology = device_topology
        self.__device_enum = device_enum
        self.__eq = None

    @property
    def channel_num(self):
        return self.__channel_num

    @property
    def eq(self):
        if self.__eq is None:
            self.__eq = Equalizer(
                self.__channel_num,
                self.__device,
                self.__device_topology,
                self.__device_enum
            )

        return self.__eq

    @property
    def bass_boost(self):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return False

        ksprop = KSNODEPROPERTY_AUDIO_CHANNEL()
        ksprop.NodeProperty.Property.Set = KSPROPSETID_Audio
        ksprop.NodeProperty.Property.Id = KSPROPERTY_AUDIO_BASS_BOOST
        ksprop.NodeProperty.Property.Flags = (
            KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY
        )
        ksprop.Channel = self.__channel_num

        for subunit in self.__device.subunits:
            part = subunit.part
            if part.sub_type != KSNODETYPE_TONE:
                continue

            local_id = part.local_id
            ksprop.NodeProperty.NodeId = local_id & PARTID_MASK
            bValue = BOOL()
            valueSize = ULONG()
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            if valueSize.value == 0:
                continue

            return bool(bValue)

        return False

    @bass_boost.setter
    def bass_boost(self, value):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return

        ksprop = KSNODEPROPERTY_AUDIO_CHANNEL()
        ksprop.NodeProperty.Property.Set = KSPROPSETID_Audio
        ksprop.NodeProperty.Property.Id = KSPROPERTY_AUDIO_BASS_BOOST
        ksprop.NodeProperty.Property.Flags = (
            KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY
        )
        ksprop.Channel = self.__channel_num

        for subunit in self.__device.subunits:
            part = subunit.part

            if part.sub_type != KSNODETYPE_TONE:
                continue

            local_id = part.local_id
            ksprop.NodeProperty.NodeId = local_id & PARTID_MASK
            bValue = BOOL(value)
            valueSize = ULONG(ctypes.sizeof(bValue))
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            break

    @property
    def loudness(self):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return False

        ksprop = KSNODEPROPERTY_AUDIO_CHANNEL()
        ksprop.NodeProperty.Property.Set = KSPROPSETID_Audio
        ksprop.NodeProperty.Property.Id = KSPROPERTY_AUDIO_LOUDNESS
        ksprop.NodeProperty.Property.Flags = (
            KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY
        )
        ksprop.Channel = self.__channel_num
        for subunit in self.__device.subunits:
            part = subunit.part
            if part.sub_type != KSNODETYPE_LOUDNESS:
                continue

            local_id = part.local_id
            ksprop.NodeProperty.NodeId = local_id & PARTID_MASK
            bValue = BOOL()
            valueSize = ULONG()
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue
                
            if valueSize.value == 0:
                continue

            return bool(bValue)

        return False

    @loudness.setter
    def loudness(self, value):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return

        ksprop = KSNODEPROPERTY_AUDIO_CHANNEL()
        ksprop.NodeProperty.Property.Set = KSPROPSETID_Audio
        ksprop.NodeProperty.Property.Id = KSPROPERTY_AUDIO_LOUDNESS
        ksprop.NodeProperty.Property.Flags = (
            KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY
        )
        ksprop.Channel = self.__channel_num

        for subunit in self.__device.subunits:
            part = subunit.part

            if part.sub_type != KSNODETYPE_LOUDNESS:
                continue

            local_id = part.local_id
            ksprop.NodeProperty.NodeId = local_id & PARTID_MASK
            bValue = BOOL(value)
            valueSize = ULONG(ctypes.sizeof(bValue))
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            break

    @property
    def automatic_gain_control(self):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return False

        ksprop = KSNODEPROPERTY_AUDIO_CHANNEL()
        ksprop.NodeProperty.Property.Set = KSPROPSETID_Audio
        ksprop.NodeProperty.Property.Id = KSPROPERTY_AUDIO_AGC
        ksprop.NodeProperty.Property.Flags = (
            KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY
        )
        ksprop.Channel = self.__channel_num
        for subunit in self.__device.subunits:
            part = subunit.part
            if part.sub_type != KSNODETYPE_AGC:
                continue

            local_id = part.local_id
            ksprop.NodeProperty.NodeId = local_id & PARTID_MASK
            bValue = BOOL()
            valueSize = ULONG()
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            if valueSize.value == 0:
                continue

            return bool(bValue)

        return False

    @automatic_gain_control.setter
    def automatic_gain_control(self, value):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return

        ksprop = KSNODEPROPERTY_AUDIO_CHANNEL()
        ksprop.NodeProperty.Property.Set = KSPROPSETID_Audio
        ksprop.NodeProperty.Property.Id = KSPROPERTY_AUDIO_AGC
        ksprop.NodeProperty.Property.Flags = (
            KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY
        )
        ksprop.Channel = self.__channel_num

        for subunit in self.__device.subunits:
            part = subunit.part

            if part.sub_type != KSNODETYPE_AGC:
                continue

            local_id = part.local_id
            ksprop.NodeProperty.NodeId = local_id & PARTID_MASK
            bValue = BOOL(value)
            valueSize = ULONG(ctypes.sizeof(bValue))
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            break

    @property
    def bass(self):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return

        ksprop = KSNODEPROPERTY_AUDIO_CHANNEL()
        ksprop.NodeProperty.Property.Set = KSPROPSETID_Audio
        ksprop.NodeProperty.Property.Id = KSPROPERTY_AUDIO_BASS
        ksprop.NodeProperty.Property.Flags = (
            KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY
        )
        ksprop.Channel = self.__channel_num

        for subunit in self.__device.subunits:
            part = subunit.part

            if part.sub_type != KSNODETYPE_AGC:
                continue

            local_id = part.local_id
            ksprop.NodeProperty.NodeId = local_id & PARTID_MASK
            bValue = LONG()
            valueSize = ULONG()
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            if valueSize.value == 0:
                continue

            return bValue.value

    @bass.setter
    def bass(self, value):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return

        ksprop = KSNODEPROPERTY_AUDIO_CHANNEL()
        ksprop.NodeProperty.Property.Set = KSPROPSETID_Audio
        ksprop.NodeProperty.Property.Id = KSPROPERTY_AUDIO_BASS
        ksprop.NodeProperty.Property.Flags = (
            KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY
        )
        ksprop.Channel = self.__channel_num

        for subunit in self.__device.subunits:
            part = subunit.part

            if part.sub_type != KSNODETYPE_TONE:
                continue

            local_id = part.local_id
            ksprop.NodeProperty.NodeId = local_id & PARTID_MASK
            bValue = LONG(value)
            valueSize = ULONG(ctypes.sizeof(bValue))
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            break

    @property
    def mid(self):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return

        ksprop = KSNODEPROPERTY_AUDIO_CHANNEL()
        ksprop.NodeProperty.Property.Set = KSPROPSETID_Audio
        ksprop.NodeProperty.Property.Id = KSPROPERTY_AUDIO_MID
        ksprop.NodeProperty.Property.Flags = (
            KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY
        )
        ksprop.Channel = self.__channel_num

        for subunit in self.__device.subunits:
            part = subunit.part

            if part.sub_type != KSNODETYPE_TONE:
                continue

            local_id = part.local_id
            ksprop.NodeProperty.NodeId = local_id & PARTID_MASK
            bValue = LONG()
            valueSize = ULONG()
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            if valueSize.value == 0:
                continue

            return bValue.value

    @mid.setter
    def mid(self, value):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return

        ksprop = KSNODEPROPERTY_AUDIO_CHANNEL()
        ksprop.NodeProperty.Property.Set = KSPROPSETID_Audio
        ksprop.NodeProperty.Property.Id = KSPROPERTY_AUDIO_MID
        ksprop.NodeProperty.Property.Flags = (
            KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY
        )
        ksprop.Channel = self.__channel_num

        for subunit in self.__device.subunits:
            part = subunit.part

            if part.sub_type != KSNODETYPE_TONE:
                continue

            local_id = part.local_id
            ksprop.NodeProperty.NodeId = local_id & PARTID_MASK
            bValue = LONG(value)
            valueSize = ULONG(ctypes.sizeof(bValue))
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            break

    @property
    def treble(self):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return

        ksprop = KSNODEPROPERTY_AUDIO_CHANNEL()
        ksprop.NodeProperty.Property.Set = KSPROPSETID_Audio
        ksprop.NodeProperty.Property.Id = KSPROPERTY_AUDIO_TREBLE
        ksprop.NodeProperty.Property.Flags = (
            KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY
        )
        ksprop.Channel = self.__channel_num
        for subunit in self.__device.subunits:
            part = subunit.part
            if part.sub_type != KSNODETYPE_TONE:
                continue

            local_id = part.local_id
            ksprop.NodeProperty.NodeId = local_id & PARTID_MASK
            bValue = LONG()
            valueSize = ULONG()
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            if valueSize.value == 0:
                continue

            return bValue.value

    @treble.setter
    def treble(self, value):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return

        ksprop = KSNODEPROPERTY_AUDIO_CHANNEL()
        ksprop.NodeProperty.Property.Set = KSPROPSETID_Audio
        ksprop.NodeProperty.Property.Id = KSPROPERTY_AUDIO_TREBLE
        ksprop.NodeProperty.Property.Flags = (
            KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY
        )
        ksprop.Channel = self.__channel_num

        for subunit in self.__device.subunits:
            part = subunit.part

            if part.sub_type != KSNODETYPE_TONE:
                continue

            local_id = part.local_id
            ksprop.NodeProperty.NodeId = local_id & PARTID_MASK
            bValue = LONG(value)
            valueSize = ULONG(ctypes.sizeof(bValue))
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            break


class Device(object):

    def __init__(self, device_enum, device_topology):
        self.__device_enum = device_enum
        self.__device_topology = device_topology(device=self)
        self.__endpoints = {}

        for _ in self:
            continue

    @property
    def connectors(self):
        res = []
        pCount = self.__device_topology.GetConnectorCount()
        for i in range(pCount):
            # noinspection PyTypeChecker
            connector = POINTER(IConnector)()
            self.__device_topology.GetConnector(i, ctypes.byref(connector))
            res.append(connector(endpoint=self))

        return res

    @property
    def subunits(self):
        res = []
        pCount = self.__device_topology.GetSubunitCount()

        for i in range(pCount):
            # noinspection PyTypeChecker
            subunit = POINTER(ISubunit)()
            self.__device_topology.GetSubunit(i, ctypes.byref(subunit))
            res.append(subunit(endpoint=self))

        return res

    @property
    def name(self):
        for endpoint in self:
            return endpoint.device_name

    @property
    def id(self):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return device_id.rsplit('\\', 1)[0]

    def __iter__(self):
        endpoint_ids = []
        id_ = self.id
        # noinspection PyTypeChecker
        endpoint_enum = POINTER(IMMDeviceCollection)()

        self.__device_enum.EnumAudioEndpoints(
            EDataFlow.eAll,
            DEVICE_STATE_MASK_ALL,
            ctypes.byref(endpoint_enum)
        )

        for endpoint in endpoint_enum:
            pDevTopoEndpt = endpoint.activate(IDeviceTopology)
            pConnEndpt = pDevTopoEndpt.connectors[0]
            pConnHWDev = pConnEndpt.connected_to
            if not pConnHWDev:
                continue

            pPartConn = pConnHWDev.part
            ppDevTopo = pPartConn.device_topology
            device_id = ppDevTopo.device_id

            if device_id != id_:
                continue

            endpoint_id = endpoint.id
            if endpoint_id in endpoint_ids:
                continue

            endpoint_ids.append(endpoint_id)

            if endpoint_id not in self.__endpoints:
                self.__endpoints[endpoint_id] = endpoint(self, pDevTopoEndpt)

        for id_ in list(self.__endpoints.keys()):
            if id_ not in endpoint_ids:
                del self.__endpoints[id_]

        for endpoint in list(self.__endpoints.values())[:]:
            yield endpoint


class IMMDevice(comtypes.IUnknown):
    """
    Main entry point for an audio endpoint
    """
    _case_insensitive_ = False
    _iid_ = IID_IMMDevice
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'Activate',
            (['in'], REFIID, 'iid'),
            (['in'], DWORD, 'dwClsCtx'),
            (['in'], PPROPVARIANT, 'pActivationParams', None),
            (['out'], POINTER(LPVOID), 'ppInterface')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'OpenPropertyStore',
            (['in'], DWORD, 'stgmAccess'),
            (['out'], POINTER(PIPropertyStore), 'ppProperties'),
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetId',
            (['out'], POINTER(LPWSTR), 'ppstrId')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetState',
            (['out'], LPDWORD, 'pdwState')
        )
    )

    def __init__(self):
        self.__chorus = None
        self.__reverb = None
        super(IMMDevice, self).__init__()

    @property
    def audio_channels(self):
        channel_config = self.channel_config
        if channel_config is None:
            return []

        res = []

        for i in range(channel_config.num_channels):
            audio_channel = AudioChannel(
                self.__device,
                i,
                self.__device_topology,
                self.__device_enum
            )
            res.append(audio_channel)

        return res

    @property
    def chorus(self):
        if self.__chorus is None:
            self.__chorus = Chorus(
                self,
                self.__device_topology,
                self.__device_enum
            )

        return self.__chorus

    @property
    def reverb(self):
        if self.__reverb is None:
            self.__reverb = Reverb(
                self,
                self.__device_topology,
                self.__device_enum
            )

        return self.__reverb

    @property
    def three_d(self):
        """
        DS3DALG_DEFAULT
        DS3DALG_NO_VIRTUALIZATION
        DS3DALG_HRTF_FULL
        DS3DALG_HRTF_LIGHT
        """
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return

        ksprop = KSNODEPROPERTY()
        ksprop.Property.Set = KSPROPSETID_Audio
        ksprop.Property.Id = KSPROPERTY_AUDIO_WIDENESS
        ksprop.Property.Flags = KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY

        for subunit in self.subunits:
            part = subunit.part
            if part.sub_type != KSNODETYPE_3D_EFFECTS:
                continue

            local_id = part.local_id
            ksprop.NodeId = local_id & PARTID_MASK
            bValue = GUID()
            valueSize = ULONG()
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            if valueSize.value == 0:
                continue

            if bValue == DS3DALG_DEFAULT:
                return DS3DALG_DEFAULT
            if bValue == DS3DALG_NO_VIRTUALIZATION:
                return DS3DALG_NO_VIRTUALIZATION
            if bValue == DS3DALG_HRTF_FULL:
                return DS3DALG_HRTF_FULL
            if bValue == DS3DALG_HRTF_LIGHT:
                return DS3DALG_HRTF_LIGHT

            break
        return

    @three_d.setter
    def three_d(self, value):
        """
        DS3DALG_DEFAULT
        DS3DALG_NO_VIRTUALIZATION
        DS3DALG_HRTF_FULL
        DS3DALG_HRTF_LIGHT
        """
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return

        ksprop = KSNODEPROPERTY()
        ksprop.Property.Set = KSPROPSETID_Audio
        ksprop.Property.Id = KSPROPERTY_AUDIO_WIDENESS
        ksprop.Property.Flags = KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY

        for subunit in self.subunits:
            part = subunit.part

            if part.sub_type != KSNODETYPE_3D_EFFECTS:
                continue

            local_id = part.local_id
            ksprop.NodeId = local_id & PARTID_MASK
            bValue = value
            valueSize = ULONG(ctypes.sizeof(bValue))
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            break

    @property
    def acoustic_echo_cancel_mode(self):
        """
        AEC_MODE_PASS_THROUGH
        AEC_MODE_HALF_DUPLEX
        AEC_MODE_FULL_DUPLEX
        """
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return False

        ksprop = KSNODEPROPERTY()
        ksprop.Property.Set = KSPROPSETID_Acoustic_Echo_Cancel
        ksprop.Property.Id = KSPROPERTY_AEC_MODE
        ksprop.Property.Flags = KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY

        for subunit in self.subunits:
            part = subunit.part
            if part.sub_type != KSNODETYPE_ACOUSTIC_ECHO_CANCEL:
                continue

            local_id = part.local_id
            ksprop.NodeId = local_id & PARTID_MASK
            bValue = ULONG()
            valueSize = ULONG()
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            if valueSize.value == 0:
                continue

            if bValue.value == AEC_MODE_PASS_THROUGH:
                return AEC_MODE_PASS_THROUGH

            if bValue.value == AEC_MODE_HALF_DUPLEX:
                return AEC_MODE_HALF_DUPLEX

            if bValue.value == AEC_MODE_FULL_DUPLEX:
                return AEC_MODE_FULL_DUPLEX

        return False

    @acoustic_echo_cancel_mode.setter
    def acoustic_echo_cancel_mode(self, value):
        """
        AEC_MODE_PASS_THROUGH
        AEC_MODE_HALF_DUPLEX
        AEC_MODE_FULL_DUPLEX
        """
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return

        ksprop = KSNODEPROPERTY()
        ksprop.Property.Set = KSPROPSETID_Acoustic_Echo_Cancel
        ksprop.Property.Id = KSPROPERTY_AEC_MODE
        ksprop.Property.Flags = KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY

        for subunit in self.subunits:
            part = subunit.part

            if part.sub_type != KSNODETYPE_ACOUSTIC_ECHO_CANCEL:
                continue

            local_id = part.local_id
            ksprop.NodeId = local_id & PARTID_MASK
            bValue = ULONG(value)
            valueSize = ULONG(ctypes.sizeof(bValue))
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            break

    @property
    def wideness(self):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return False

        ksprop = KSNODEPROPERTY()
        ksprop.Property.Set = KSPROPSETID_Audio
        ksprop.Property.Id = KSPROPERTY_AUDIO_WIDENESS
        ksprop.Property.Flags = KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY

        for subunit in self.subunits:
            part = subunit.part
            if part.sub_type != KSNODETYPE_STEREO_WIDE:
                continue

            local_id = part.local_id
            ksprop.NodeId = local_id & PARTID_MASK
            bValue = ULONG()
            valueSize = ULONG()
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            if valueSize.value == 0:
                continue

            return bool(bValue)

        return False

    @wideness.setter
    def wideness(self, value):
        data = self.__device_topology.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)

        imm_device = self.__device_enum.GetDevice(device_id)
        iks_control = imm_device.activate(IKsControl)
        if not iks_control:
            return

        ksprop = KSNODEPROPERTY()
        ksprop.Property.Set = KSPROPSETID_Audio
        ksprop.Property.Id = KSPROPERTY_AUDIO_WIDENESS
        ksprop.Property.Flags = KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY

        for subunit in self.subunits:
            part = subunit.part

            if part.sub_type != KSNODETYPE_STEREO_WIDE:
                continue

            local_id = part.local_id
            ksprop.NodeId = local_id & PARTID_MASK
            bValue = ULONG(value)
            valueSize = ULONG(ctypes.sizeof(bValue))
            try:
                iks_control.KsProperty(
                    ctypes.byref(ksprop),
                    ctypes.sizeof(ksprop),
                    ctypes.byref(bValue),
                    ctypes.sizeof(bValue),
                    ctypes.byref(valueSize)
                )
            except comtypes.COMError:
                continue

            break

    @property
    def device_name(self):
        """
        Name of the device this endpoint belongs to
        """
        return self.get_property(PKEY_DeviceInterface_FriendlyName)

    def get_properties(self):
        pStore = self.OpenPropertyStore(STGM_READ)

        count = pStore.GetCount()

        for i in range(count):
            p_key = pStore.GetAt(i)
            print(p_key)
            try:
                prop_var = self.get_property(p_key, True)
                print('prop VT:', prop_var.vt)
            except:
                print('no read permissions')

            print()

    def get_property(self, key, propvar=False):
        """
        Internal Use
        """
        # noinspection PyUnresolvedReferences
        pStore = self.OpenPropertyStore(STGM_READ)
        try:
            return pStore.GetValue(key, propvar)
        except comtypes.COMError:
            raise AttributeError

    def set_property(self, key, value, vt=None):
        """
        Internal use
        """
        # noinspection PyUnresolvedReferences
        pStore = self.OpenPropertyStore(STGM_WRITE)

        try:
            return pStore.SetValue(key, value, vt)
        except comtypes.COMError:
            raise AttributeError

    def activate(self, cls):
        """
        Internal use
        """
        try:
            # noinspection PyUnresolvedReferences
            return ctypes.cast(
                self.Activate(
                    cls._iid_,
                    comtypes.CLSCTX_INPROC_SERVER,
                    None
                ),
                POINTER(cls)
            )

        except comtypes.COMError:
            pass

    @property
    def state(self) -> int:
        # noinspection PyUnresolvedReferences
        state = self.GetState()
        if state == DEVICE_STATE_ACTIVE:
            return DEVICE_STATE_ACTIVE

        if state == DEVICE_STATE_DISABLED:
            return DEVICE_STATE_DISABLED

        if state == DEVICE_STATE_NOTPRESENT:
            return DEVICE_STATE_NOTPRESENT

        if state == DEVICE_STATE_UNPLUGGED:
            return DEVICE_STATE_UNPLUGGED

    @property
    def id(self):
        # noinspection PyUnresolvedReferences
        data = self.GetId()
        id_ = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return id_

    @property
    def audio_client(self):
        # noinspection PyTypeChecker
        audio_client = POINTER(IAudioClient)()
        self.activate(audio_client)
        return audio_client(self)

    @property
    def data_flow(self) -> EDataFlow:
        """
        The direction of the audio (in/out)

        one of the `EDataFlow` attributes
        """
        endpoint = self.QueryInterface(IMMEndpoint)
        if endpoint:
            return EDataFlow.get(endpoint.GetDataFlow().value)

    @property
    def name(self):
        """
        Endpoints friendly name
        """
        val = self.get_property(PKEY_Device_FriendlyName)
        return utils.convert_to_string(val)

    @property
    def description(self):
        """
        Endpoints description
        """
        val = self.get_property(PKEY_Device_DeviceDesc)
        return utils.convert_to_string(val)

    @property
    def form_factor(self) -> EndpointFormFactor:
        """
        The endpoint  form factor

        One of `EndpointFormFactor` attributes
        """
        form_factor = self.get_property(PKEY_AudioEndpoint_FormFactor)
        return EndpointFormFactor.get(form_factor)

    @property
    def type(self) -> NodeTypeGUID:
        sub_type = self.get_property(PKEY_AudioEndpoint_JackSubType)
        sub_type_guid = str(sub_type)
        value = NodeTypeGUID.get(sub_type_guid)
        if value is None:
            return NodeTypeGUID('Unknown', GUID_NULL)
        return value

    @property
    def full_range_speakers(self) -> AudioSpeakers:
        return AudioSpeakers(
            self.get_property(PKEY_AudioEndpoint_FullRangeSpeakers)
        )

    @property
    def guid(self) -> GUID:
        return self.get_property(PKEY_AudioEndpoint_GUID)

    @property
    def physical_speakers(self) -> AudioSpeakers:
        return AudioSpeakers(
            self.get_property(PKEY_AudioEndpoint_PhysicalSpeakers)
        )

    @property
    def audio_enhancements_enabled(self) -> bool:
        prop_var = self.get_property(
            PKEY_AudioEndpoint_Disable_SysFx,
            propvar=True
        )

        return prop_var.ulVal == ENDPOINT_SYSFX_ENABLED

    @audio_enhancements_enabled.setter
    def audio_enhancements_enabled(self, value: bool):
        from .propidl import VT_UI4

        if value:
            value = ENDPOINT_SYSFX_ENABLED
        else:
            value = ENDPOINT_SYSFX_DISABLED

        self.set_property(PKEY_AudioEndpoint_Disable_SysFx, value, VT_UI4)

    @property
    def hdcp_capable(self) -> bool:
        js = self.jack_sink_information
        if js is None:
            return False
        return js.hdcp_capable

    @property
    def ai_capable(self) -> bool:
        js = self.jack_sink_information
        if js is None:
            return False

        return js.ai_capable

    @property
    def connector_type(self) -> ConnectorType:
        connector = self.connector.connected_to
        return connector.type

    @property
    def connector_name(self) -> str:
        connector = self.connector.connected_to
        part = connector.part

        return part.name

    @property
    def connector_subtype(self) -> KSNODETYPE:
        connector = self.connector.connected_to
        part = connector.part

        return part.sub_type

    @property
    def connector_location(self) -> str:
        jd = self.jack_descriptions
        if not jd:
            return 'Unknown'

        jd = jd[0]

        return jd.location

    @property
    def connector_style(self) -> EPcxConnectionType:
        js = self.jack_sink_information
        if js is not None:
            return js.connection_type

        jd = self.jack_descriptions
        if not jd:
            return ENUM_VALUE(-1, 'Unknown')

        jd = jd[0]

        return jd.connection_type

    @property
    def presence_detection(self) -> bool:
        jd = self.jack_descriptions
        if not jd:
            return False

        jd = jd[0]

        return jd.presence_detection

    @property
    def connector_color(self) -> tuple:
        jd = self.jack_descriptions
        if not jd:
            if 'green' in self.connector_name.lower():
                return 0, 255, 0, 255
            if 'pink' in self.connector_name.lower():
                return 255, 128, 192, 255
            if 'blue' in self.connector_name.lower():
                return 0, 0, 255, 255

            return 0, 0, 0, 0

        jd = jd[0]

        return jd.color + (255,)

    @property
    def is_connected(self) -> bool:
        jd = self.jack_descriptions
        if not jd:
            return False

        jd = jd[0]

        return jd.is_connected

    @property
    def jack_descriptions(self) -> list:
        conn_to = self.connector.connected_to
        if conn_to is None:
            return []

        part = conn_to.part
        jack_description = part.activate(IKsJackDescription)
        if not jack_description:
            return []

        jack_description2 = part.activate(IKsJackDescription2)
        if not jack_description2:
            jack_description2 = None

        jds = list(jack_description)

        if jack_description2 is None:
            jd2s = [None] * len(jds)
        else:
            jd2s = list(jack_description2)

        res = []

        for jd1, jd2 in zip(jds, jd2s):
            res.append(JackDescription(jd1, jd2))

        return res

    @property
    def jack_sink_information(self) -> KSJACK_SINK_INFORMATION:
        conn_to = self.connector.connected_to
        if conn_to is None:
            return KSJACK_SINK_INFORMATION()

        part = conn_to.part
        sink_information = part.activate(IKsJackSinkInformation)

        if sink_information:
            return sink_information.GetJackSinkInformation()

    @property
    def has_audio(self):
        volume = self.volume

        if volume is None:
            return -1

        # noinspection PyProtectedMember
        peak_meter = volume._peak_meter
        if peak_meter is None:
            return -1

        pfPeak = FLOAT()
        peak_meter.GetPeakValue(ctypes.byref(pfPeak))
        return pfPeak.value > 1E-08

    @property
    def channel_config(self) -> IAudioChannelConfig:
        return self.__get_interface(IAudioChannelConfig)

    @property
    def input(self) -> IAudioInputSelector:
        return self.__get_interface(IAudioInputSelector)

    @property
    def output(self) -> IAudioOutputSelector:
        return self.__get_interface(IAudioOutputSelector)

    def __get_interface(self, cls):
        # The device topology for an endpoint device always
        # contains just one connector (connector number 0).
        outgoing = True
        pConnFrom = self.connector

        # Outer loop: Each iteration traverses the data path
        # through a device topology starting at the input
        # connector and ending at the output connector.
        while True:

            #  Does this connector connect to another device?
            if not pConnFrom.is_connected:
                # This is the end of the data path that
                # stretches from the endpoint device to the
                # system bus or external bus. Verify that
                # the connection type is Software_IO.
                if pConnFrom.type == ConnectorType.Software_IO:
                    break

            # Get the connector in the next device topology,
            # which lies on the other side of the connection.
            pConnTo = pConnFrom.connected_to

            # Get the connector's IPart interface.
            pPartPrev = pConnTo.part

            # Inner loop: Each iteration traverses one link in a
            # device topology and looks for input multiplexers.
            while True:
                # Follow downstream link to next part.
                if outgoing:
                    pParts = pPartPrev.outgoing
                else:
                    pParts = pPartPrev.incoming

                if not pParts:
                    if outgoing:
                        outgoing = False
                        pParts = pPartPrev.incoming
                        if not pParts:
                            return
                    else:
                        return

                #     pParts = pPartPrev.incoming
                #
                # if not pParts:
                #     return

                pPartNext = list(pParts)[0]
                parttype = pPartNext.type

                # Failure of the following call means only that
                # the part is not a MUX (input selector).
                for p in pParts:
                    for i in p:
                        if isinstance(i, cls):
                            return i

                    interface = p.activate(cls)

                    if interface:
                        return interface

                if parttype == PartType.Connector:
                    # We've reached the output connector that
                    # lies at the end of this device topology.
                    pConnFrom = pPartNext.connector
                    break

                pPartPrev = pPartNext

        #
        # while True:
        #     try:
        #         conn_from = conn_from.connected_to
        #     except comtypes.COMError:
        #         return None
        #
        #     part = conn_from.part
        #     device_topology = part.device_topology
        #     for subunit in device_topology.subunits:
        #         part = subunit.part
        #
        #
        #     if conn_from.type == ConnectorType.Software_IO:
        #         return None
        #
        #     if not conn_from.is_connected:
        #         return None

    @property
    def volume(self) -> IAudioEndpointVolumeEx:
        return self.__volume

    def __call__(self, device, device_topology):
        self.__device = device
        self.__device_topology = device_topology(endpoint=self)
        self.__chorus = None
        self.__reverb = None
        self.__device_enum = device._Device__device_enum  # NOQA

        session_manager = self.activate(IAudioSessionManager2)
        if not session_manager:
            session_manager = self.activate(IAudioSessionManager)

        if session_manager:
            # noinspection PyCallingNonCallable
            session_manager = session_manager(endpoint=self)
        else:
            session_manager = None

        self.__session_manager = session_manager

        endpoint_volume = self.activate(IAudioEndpointVolumeEx)

        if not endpoint_volume:
            endpoint_volume = self.activate(IAudioEndpointVolume)

        if endpoint_volume:
            # noinspection PyCallingNonCallable
            endpoint_volume = endpoint_volume(endpoint=self)
        else:
            endpoint_volume = None

        self.__volume = endpoint_volume

        return self

    @property
    def connector(self) -> IConnector:
        # noinspection PyTypeChecker
        connector = POINTER(IConnector)()
        self.__device_topology.GetConnector(0, ctypes.byref(connector))
        return connector(endpoint=self)

    @property
    def subunits(self) -> list:
        res = []
        pCount = self.__device_topology.GetSubunitCount()

        for i in range(pCount):
            # noinspection PyTypeChecker
            subunit = POINTER(ISubunit)()
            self.__device_topology.GetSubunit(i, ctypes.byref(subunit))
            res.append(subunit(endpoint=self))

        return res

    @property
    def device(self) -> Device:
        return self.__device

    def set_default(self, role):
        policy_config = comtypes.CoCreateInstance(
            policyconfig.CLSID_PolicyConfigClient,
            policyconfig.IPolicyConfig,
            comtypes.CLSCTX_ALL
        )

        policy_config.SetDefaultEndpoint(self.id, ERole.get(role))

    @property
    def is_default(self) -> bool:
        for role in (
            ERole.eMultimedia,
            ERole.eConsole,
            ERole.eCommunications
        ):
            if self.is_default_endpoint(role):
                return True

        return False

    def is_default_endpoint(
        self,
        role: Union[ERole, str] = ERole.eMultimedia
    ) -> bool:

        return IMMDeviceEnumerator.default_audio_endpoint(
            self.data_flow,
            ERole.get(role)
        ) == self

    def __iter__(self):
        """
        Sessions
        """
        if self.__session_manager is not None:
            for session in self.__session_manager:
                yield session


# noinspection PyTypeChecker
PIMMDevice = POINTER(IMMDevice)

from .audiopolicy import (
    IAudioSessionManager,
    IAudioSessionManager2,
)  # NOQA


class IMMDeviceActivator(comtypes.IUnknown):
    _iid_ = IID_IMMDeviceActivator
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'Activate',
            (['in'], REFIID, 'iid'),
            (['in'], PIMMDevice, 'pDevice'),
            (['in'], PPROPVARIANT, 'pActivationParams'),
            (['in'], POINTER(LPVOID), 'ppInterface'),
        ),
    )


# noinspection PyTypeChecker
PIMMDeviceActivator = POINTER(IMMDeviceActivator)


class IMMDeviceCollection(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IMMDeviceCollection
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetCount',
            (['out'], UINT_PTR, 'pcDevices')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'Item',
            (['in'], UINT, 'nDevice'),
            (['in'], POINTER(PIMMDevice), 'ppDevice')
        )
    )

    def __iter__(self):
        # noinspection PyUnresolvedReferences
        count = self.GetCount()
        for i in range(count):
            # noinspection PyTypeChecker
            item = POINTER(IMMDevice)()
            # noinspection PyUnresolvedReferences
            self.Item(UINT(i), ctypes.byref(item))
            yield item


# noinspection PyTypeChecker
PIMMDeviceCollection = ctypes.POINTER(IMMDeviceCollection)


class _IMMDeviceEnumerator(comtypes.IUnknown):
    _devices = {}
    _case_insensitive_ = False
    _iid_ = IID_IMMDeviceEnumerator
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'EnumAudioEndpoints',
            (['in'], EDataFlow, 'dataFlow'),
            (['in'], DWORD, 'dwStateMask'),
            (['in'], POINTER(PIMMDeviceCollection), 'ppDevices')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetDefaultAudioEndpoint',
            (['in'], EDataFlow, 'dataFlow'),
            (['in'], ERole, 'role'),
            (['out'], POINTER(PIMMDevice), 'ppDevices')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetDevice',
            (['in'], LPCWSTR, 'pwstrId'),
            (['out'], POINTER(PIMMDevice), 'ppDevices')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'RegisterEndpointNotificationCallback',
            (['in'], PIMMNotificationClient, 'pClient'),
        ),
        COMMETHOD(
            [],
            HRESULT,
            'UnregisterEndpointNotificationCallback',
            (['in'], PIMMNotificationClient, 'pClient'),
        )
    )

    def __iter__(self):
        device_ids = []
        # noinspection PyTypeChecker
        endpoint_enum = POINTER(IMMDeviceCollection)()
        # noinspection PyUnresolvedReferences
        self.EnumAudioEndpoints(
            EDataFlow.eAll,
            DEVICE_STATE_MASK_ALL,
            ctypes.byref(endpoint_enum)
        )

        for endpoint in endpoint_enum:
            pDevTopoEndpt = endpoint.activate(IDeviceTopology)
            pConnEndpt = pDevTopoEndpt.connectors[0]
            pConnHWDev = pConnEndpt.connected_to
            if not pConnHWDev:
                continue
            
            pPartConn = pConnHWDev.part
            ppDevTopo = pPartConn.device_topology
            device_id = ppDevTopo.device_id

            if device_id not in device_ids:
                device_ids.append(device_id)

                if device_id not in self._devices:
                    self._devices[device_id] = Device(self, ppDevTopo)

        for id_ in list(self._devices.keys()):
            if id_ not in device_ids:
                del self._devices[id_]

        for device in list(self._devices.values())[:]:
            yield device


class IMMDeviceEnumerator(object):
    _instance = None

    def __init__(self):
        if IMMDeviceEnumerator._instance is None:
            IMMDeviceEnumerator._instance = self

            self._device_enum = comtypes.CoCreateInstance(
                CLSID_MMDeviceEnumerator,
                _IMMDeviceEnumerator,
                comtypes.CLSCTX_ALL
            )

            self.__notification_client = IMMNotificationClient(self)
            self._device_enum.RegisterEndpointNotificationCallback(
                self.__notification_client
            )

        else:
            self.__dict__.update(IMMDeviceEnumerator._instance.__dict__)

    def __del__(self):
        if self.__notification_client is not None:
            self._device_enum.UnregisterEndpointNotificationCallback(
                self.__notification_client
            )
            self.__notification_client = None

        IMMDeviceEnumerator._instance = None

    def __iter__(self):
        for device in self._device_enum:
            yield device

    @classmethod
    def default_audio_endpoint(cls, data_flow, role):
        self = cls._instance
        # noinspection PyTypeChecker
        endp = self._device_enum.GetDefaultAudioEndpoint(data_flow, role)

        id_ = endp.id
        for device in self:
            for endpoint in device:
                if endpoint.id == id_:
                    return endpoint


class __MIDL___MIDL_itf_mmdeviceapi_0000_0008_0002(ENUM):
    AUDIO_SYSTEMEFFECTS_PROPERTYSTORE_TYPE_DEFAULT = 0
    AUDIO_SYSTEMEFFECTS_PROPERTYSTORE_TYPE_USER = 1
    AUDIO_SYSTEMEFFECTS_PROPERTYSTORE_TYPE_VOLATILE = 2
    AUDIO_SYSTEMEFFECTS_PROPERTYSTORE_TYPE_ENUM_COUNT = (
        AUDIO_SYSTEMEFFECTS_PROPERTYSTORE_TYPE_VOLATILE + 1
    )


AUDIO_SYSTEMEFFECTS_PROPERTYSTORE_TYPE = (
    __MIDL___MIDL_itf_mmdeviceapi_0000_0008_0002
)


IID_IAudioSystemEffectsPropertyChangeNotificationClient = IID(
    '{20049D40-56D5-400E-A2EF-385599FEED49}'
)


class IAudioSystemEffectsPropertyChangeNotificationClient(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioSystemEffectsPropertyChangeNotificationClient
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'OnPropertyChanged',
            (['in'], AUDIO_SYSTEMEFFECTS_PROPERTYSTORE_TYPE, 'type'),
            (['in'], PROPERTYKEY, 'key')
        ),
    )


IID_IAudioSystemEffectsPropertyStore = IID(
    '{302AE7F9-D7E0-43E4-971B-1F8293613D2A}'
)


class IAudioSystemEffectsPropertyStore(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioSystemEffectsPropertyStore
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'OpenDefaultPropertyStore',
            (['in'], DWORD, 'stgmAccess'),
            (['out'], POINTER(PIPropertyStore), 'propStore')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'OpenUserPropertyStore',
            (['in'], DWORD, 'stgmAccess'),
            (['out'], POINTER(PIPropertyStore), 'propStore')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'OpenVolatilePropertyStore',
            (['in'], DWORD, 'stgmAccess'),
            (['out'], POINTER(PIPropertyStore), 'propStore')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'ResetUserPropertyStore'
        ),
        COMMETHOD(
            [],
            HRESULT,
            'ResetVolatilePropertyStore'
        ),
        COMMETHOD(
            [],
            HRESULT,
            'RegisterPropertyChangeNotification',
            (['in'],
             POINTER(IAudioSystemEffectsPropertyChangeNotificationClient),
             'callback')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'UnregisterPropertyChangeNotification',
            (['in'],
             POINTER(IAudioSystemEffectsPropertyChangeNotificationClient),
             'callback')
        )
    )


class MMDeviceAPILib(object):
    name = u'MMDeviceAPILib'
    _reg_typelib_ = (IID_MMDeviceAPILib, 1, 0)


class MMDeviceEnumerator(CoClass):
    _reg_clsid_ = CLSID_MMDeviceEnumerator
    _idlflags_ = []
    _reg_typelib_ = (IID_MMDeviceAPILib, 1, 0)
    _com_interfaces_ = [IMMDeviceEnumerator]


class AudioExtensionParams(ctypes.Structure):
    _fields_ = [
        ('AddPageParam', LPARAM),
        ('pEndpoint', PIMMDevice),
        ('pEndpoint', PIMMDevice),
        ('pPnpInterface', PIMMDevice),
        ('pPnpDevnode', PIMMDevice)
    ]
