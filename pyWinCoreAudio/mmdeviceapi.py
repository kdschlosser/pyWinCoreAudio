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
from .propertystore import (
    PROPERTYKEY,
    PIPropertyStore,
    PPROPVARIANT
)
from .constant import (
    S_OK,
    STGM_READ,
    STGM_WRITE,
)
from .functiondiscoverykeys_devpkey import (
    PKEY_DeviceInterface_FriendlyName,
    PKEY_Device_FriendlyName,
    PKEY_Device_DeviceDesc,
    PKEY_AudioEndpoint_FormFactor,
    PKEY_AudioEndpoint_JackSubType,
    PKEY_AudioEndpoint_FullRangeSpeakers,
    PKEY_AudioEndpoint_GUID,
    PKEY_AudioEndpoint_PhysicalSpeakers,
    PKEY_AudioEndpoint_Disable_SysFx
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

ENDPOINT_SYSFX_ENABLED = 0x00000000
ENDPOINT_SYSFX_DISABLED = 0x00000001


from .policyconfig import (
    CLSID_PolicyConfigVistaClient,
    IPolicyConfigVista
)  # NOQA
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
        print('OnDeviceStateChanged')
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
                    ON_DEVICE_STATE_CHANGED.signal(device=device, new_state=state)
                    break

                for endpoint in device:
                    if endpoint.id == dev_id:
                        ON_DEVICE_STATE_CHANGED.signal(device=device, endpoint=endpoint, new_state=state)
                        break
                else:
                    continue

                break

        tw.add(_do, pwstrDeviceId, dwNewState)

        return S_OK

    def OnDeviceAdded(self, pwstrDeviceId):
        print('OnDeviceAdded')

        pwstrDeviceId = utils.convert_to_string(pwstrDeviceId)

        def _do(dev_id):
            for device in self.__device_enum:
                if device.id == dev_id:
                    ON_DEVICE_ADDED.signal(device=device)
                    break

        tw.add(_do, pwstrDeviceId)

        return S_OK

    def OnDeviceRemoved(self, pwstrDeviceId):
        print('OnDeviceRemoved')

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

        if (pwstrDefaultDeviceId, flow, role) not in self.__last_default_endpoint:
            self.__last_default_endpoint.append((pwstrDefaultDeviceId, flow, role))
            while len(self.__last_default_endpoint) > 3:
                self.__last_default_endpoint.pop(0)

            tw.add(_do, pwstrDefaultDeviceId, flow, role)

        return S_OK

    def OnPropertyValueChanged(self, pwstrDeviceId, key):
        print('OnPropertyValueChanged')

        pwstrDeviceId = utils.convert_to_string(pwstrDeviceId)

        def _do(dev_id, k):

            for device in self.__device_enum:
                if device.id == dev_id:
                    ON_DEVICE_PROPERTY_CHANGED.signal(device=device, key=k)
                    break

                for endpoint in device:
                    if endpoint.id == dev_id:
                        ON_DEVICE_PROPERTY_CHANGED.signal(device=device, endpoint=endpoint, key=k)
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

    def __init__(self, band, frequency, device, channel_num, device_topology, device_enum):
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
            ksprop.NodeProperty.Property.Flags = KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY
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
        ksprop.NodeProperty.Property.Flags = KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY
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

            ksprop.NodeProperty.Property.Flags = KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY
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
        ksprop.NodeProperty.Property.Flags = KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY
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
        ksprop.NodeProperty.Property.Flags = KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY
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
        ksprop.NodeProperty.Property.Flags = KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY
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
        ksprop.NodeProperty.Property.Flags = KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY
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
        ksprop.NodeProperty.Property.Flags = KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY
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
        ksprop.NodeProperty.Property.Flags = KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY
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
        ksprop.NodeProperty.Property.Flags = KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY
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
        ksprop.NodeProperty.Property.Flags = KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY
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
        ksprop.NodeProperty.Property.Flags = KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY
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
        ksprop.NodeProperty.Property.Flags = KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY
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
        ksprop.NodeProperty.Property.Flags = KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY
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
        ksprop.NodeProperty.Property.Flags = KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY
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
        ksprop.NodeProperty.Property.Flags = KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY
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

    @property
    def audio_channels(self):
        channel_config = self.channel_config
        if channel_config is None:
            return []

        res = []

        for i in range(channel_config.num_channels):
            audio_channel = AudioChannel(self.__device, i, self.__device_topology, self.__device_enum)
            res.append(audio_channel)

        return res

    @property
    def chorus(self):
        if self.__chorus is None:
            self.__chorus = Chorus(self, self.__device_topology, self.__device_enum)

        return self.__chorus

    @property
    def reverb(self):
        if self.__reverb is None:
            self.__reverb = Reverb(self, self.__device_topology, self.__device_enum)

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

    def get_property(self, key):
        """
        Internal Use
        """
        # noinspection PyUnresolvedReferences
        pStore = self.OpenPropertyStore(STGM_READ)
        try:
            return pStore.GetValue(key)
        except comtypes.COMError:
            raise AttributeError

    def set_property(self, key, _):
        """
        Internal use
        """
        # noinspection PyUnresolvedReferences
        pStore = self.OpenPropertyStore(STGM_WRITE)
        try:
            return pStore.GetValue(key)
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
    def system_effects(self) -> bool:
        return bool(
            self.get_property(PKEY_AudioEndpoint_Disable_SysFx)
        )

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
        self.__device_enum = device._Device__device_enum


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
            CLSID_PolicyConfigVistaClient,
            IPolicyConfigVista,
            comtypes.CLSCTX_ALL
        )

        policy_config.SetDefaultEndpoint(self.id, ERole.get(role))

    @property
    def is_default(self) -> bool:
        return IMMDeviceEnumerator.default_audio_endpoint(self.data_flow, self.data_flow) == self

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
            self._device_enum.RegisterEndpointNotificationCallback(self.__notification_client)

        else:
            self.__dict__.update(IMMDeviceEnumerator._instance.__dict__)

    def __del__(self):
        if self.__notification_client is not None:
            self._device_enum.UnregisterEndpointNotificationCallback(self.__notification_client)
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
