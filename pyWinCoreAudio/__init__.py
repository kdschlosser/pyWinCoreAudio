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


from .signal import (  # NOQA
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

from .ksmedia import (
    PKEY_AudioEndpoint_RealtekChannelsState,
    PKEY_AudioEndpoint_RealtekChannelLevelOffset,
    PKEY_AudioEndpoint_RealtekChannelDistance
)

from .audioenginebaseapo import (
    PKEY_FX_FriendlyName,
    PKEY_FX_KeywordDetector_StreamEffectClsid,
    PKEY_FX_KeywordDetector_ModeEffectClsid,
    PKEY_FX_KeywordDetector_EndpointEffectClsid,
    PKEY_FX_Offload_StreamEffectClsid,
    PKEY_FX_Offload_ModeEffectClsid,
    PKEY_CompositeFX_KeywordDetector_StreamEffectClsid,
    PKEY_CompositeFX_KeywordDetector_ModeEffectClsid,
    PKEY_CompositeFX_KeywordDetector_EndpointEffectClsid,
    PKEY_SFX_ProcessingModes_Supported_For_Streaming,
    PKEY_MFX_ProcessingModes_Supported_For_Streaming,
    PKEY_EFX_ProcessingModes_Supported_For_Streaming,
    PKEY_SFX_KeywordDetector_ProcessingModes_Supported_For_Streaming,
    PKEY_MFX_KeywordDetector_ProcessingModes_Supported_For_Streaming,
    PKEY_EFX_KeywordDetector_ProcessingModes_Supported_For_Streaming,
    PKEY_SFX_Offload_ProcessingModes_Supported_For_Streaming,
    PKEY_MFX_Offload_ProcessingModes_Supported_For_Streaming,
    PKEY_APO_SWFallback_ProcessingModes,
    PKEY_AudioEndpoint_Association
)
from .functiondiscoverykeys_devpkey import (
    PKEY_Device_FriendlyName,
    PKEY_Device_DeviceDesc,
    PKEY_DeviceInterface_FriendlyName
)
from .mmdeviceapi import (
    PKEY_AudioEndpoint_FormFactor,
    PKEY_AudioEndpoint_PhysicalSpeakers,
    PKEY_AudioEndpoint_GUID,
    PKEY_AudioEndpoint_Disable_SysFx,
    PKEY_AudioEndpoint_FullRangeSpeakers,
    PKEY_AudioEndpoint_JackSubType,
    PKEY_AudioEndpoint_ControlPanelPageProvider,
    PKEY_AudioEndpoint_Supports_EventDriven_Mode,
    PKEY_AudioEndpoint_Default_VolumeInDb,
    PKEY_AudioEngine_DeviceFormat,
    PKEY_AudioEngine_OEMFormat,
    PKEY_AudioEndpointLogo_IconEffects,
    PKEY_AudioEndpointLogo_IconPath,
    PKEY_AudioEndpointSettings_MenuText,
    PKEY_AudioEndpointSettings_LaunchContract,
    PKEY_SYSFX_Association,
    PKEY_SYSFX_PreMixClsid,
    PKEY_SYSFX_PostMixClsid,
    PKEY_SYSFX_UiClsid,
    PKEY_SYSFX_StreamEffectClsid,
    PKEY_SYSFX_ModeEffectClsid,
    PKEY_SYSFX_EndpointEffectClsid,
    PKEY_SYSFX_ChainClsid,
    PKEY_AudioEndpoint_Icon,
    PKEY_AudioDevice_EnableEndpointByDefault,
    PKEY_LFX_ProcessingModes_Supported_For_Streaming,
    PKEY_AudioEngine_PhysicalSpeaker,
    PKEY_AudioEngine_FullRangeSpeaker,
    PKEY_EFFECTNODEINFO_RENDER,
    PKEY_EFFECTNODEINFO_CAPTURE,
    PKEY_CTGUID,
    PKEY_CT_LINETOAUX_GUID,
    PKEY_FX_UiClsid,
    PKEY_ItemNameDisplay,
    PKEY_FX_PreMixClsid,
    PKEY_FX_PostMixClsid,
    PKEY_AudioDevice_NeverSetAsDefaultEndpoint,
    PKEY_Composite_SFX,
    PKEY_Composite_MFX,
    PKEY_Composite_EFX,
    PKEY_Composite_Offload_SFX,
    PKEY_Composite_Offload_MFX,
    PKEY_AudioEndpoint_ControlPanelProvider,
    PKEY_AudioEndpoint_Ext_UiClsid,
    PKEY_BYPASS_TP_EFFECTS
)
from .propkey import (
    PKEY_Audio_ChannelCount,
    PKEY_Audio_Compression,
    PKEY_Audio_EncodingBitrate,
    PKEY_Audio_Format,
    PKEY_Audio_IsVariableBitRate,
    PKEY_Audio_PeakValue,
    PKEY_Audio_SampleRate,
    PKEY_Audio_SampleSize,
    PKEY_Audio_StreamName,
    PKEY_Audio_StreamNumber,
    PKEY_Devices_AudioDevice_Microphone_IsFarField,
    PKEY_Devices_AudioDevice_Microphone_SensitivityInDbfs,
    PKEY_Devices_AudioDevice_Microphone_SensitivityInDbfs2,
    PKEY_Devices_AudioDevice_Microphone_SignalToNoiseRatioInDb,
    PKEY_Devices_AudioDevice_RawProcessingSupported,
    PKEY_Devices_AudioDevice_SpeechProcessingSupported
)


class __Module(object):

    def __init__(self):
        import sys

        mod = sys.modules[__name__]
        object.__setattr__(self, '__name__', __name__)
        object.__setattr__(self, '__doc__', mod.__doc__)
        object.__setattr__(self, '__file__', mod.__file__)
        object.__setattr__(self, '__loader__', mod.__loader__)
        object.__setattr__(self, '__package__', mod.__package__)
        object.__setattr__(self, '__path__', mod.__path__)
        object.__setattr__(self, '__spec__', mod.__spec__)
        object.__setattr__(self, '_device_enumerator', None)
        object.__setattr__(self, '__original_module__', mod)
        object.__setattr__(self, '__version__', __version__)
        object.__setattr__(self, '__author__', __author__)
        object.__setattr__(self, '__description__', __description__)
        object.__setattr__(self, '__url__', __url__)
        self.__dict__.update(mod.__dict__)
        sys.modules[__name__] = self

    def __getattr__(self, item):
        if item in self.__dict__:
            return self.__dict__[item]

        if hasattr(self.__original_module__, item):
            value = getattr(self.__original_module__, item)
            object.__setattr__(self, item, value)
            return value

        raise AttributeError(item)

    def __setattr__(self, key, value):
        raise AttributeError('access denied')

    def __iter__(self):
        if self._device_enumerator is None:
            import comtypes

            comtypes.CoInitialize()
            from .mmdeviceapi import IMMDeviceEnumerator
            object.__setattr__(self, '_device_enumerator', IMMDeviceEnumerator())

        for device in self._device_enumerator:
            yield device

    # def __del__(self):
    #     import comtypes
    #
    #     if self._device_enumerator is not None:
    #         object.__setattr__(self, '_device_enumerator', None)
    #         comtypes.CoUninitialize()

    def unload(self):
        import comtypes

        if self._device_enumerator is not None:
            object.__setattr__(self, '_device_enumerator', None)
            comtypes.CoUninitialize()


__Module()


def unload():
    pass


def __iter__(self):
    pass

