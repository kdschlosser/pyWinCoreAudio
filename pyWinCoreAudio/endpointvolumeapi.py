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
from .constant import S_OK
from .signal import ON_ENDPOINT_VOLUME_CHANGED


ENDPOINT_HARDWARE_SUPPORT_VOLUME = 0x00000001
ENDPOINT_HARDWARE_SUPPORT_MUTE = 0x00000002
ENDPOINT_HARDWARE_SUPPORT_METER = 0x00000004


IID_IAudioEndpointVolumeCallback = IID(
    '{657804FA-D6AD-4496-8A60-352752AF4F89}'
)
IID_IAudioEndpointVolumeEx = IID(
    '{66E11784-F695-4F28-A505-A7080081A78F}'
)
IID_IAudioMeterInformation = IID(
    '{C02216F6-8C67-4B5B-9D00-D008E73E0064}'
)
IID_IAudioEndpointVolume = IID(
    '{5CDF2C82-841E-4546-9722-0CF74078229A}'
)


class AUDIO_VOLUME_NOTIFICATION_DATA(ctypes.Structure):
    _fields_ = [
        ('guidEventContext', GUID),
        ('bMuted', BOOL),
        ('fMasterVolume', FLOAT),
        ('nChannels', UINT),
        ('afChannelVolumes', (FLOAT * 8))
    ]

    @property
    def guid_event_context(self):
        return self.guidEventContext

    @property
    def is_muted(self):
        return bool(self.bMuted)

    @property
    def master_volume(self):
        return self.fMasterVolume * 100.0

    @property
    def num_channels(self):
        return self.nChannels

    def __iter__(self):
        for i in range(self.num_channels):
            yield self.afChannelVolumes[i] * 100.0


PAUDIO_VOLUME_NOTIFICATION_DATA = POINTER(AUDIO_VOLUME_NOTIFICATION_DATA)


class _IAudioEndpointVolumeCallback(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioEndpointVolumeCallback
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'OnNotify',
            (['in'], PAUDIO_VOLUME_NOTIFICATION_DATA, 'pNotify')
        ),
    )


# noinspection PyTypeChecker
PIAudioEndpointVolumeCallback = POINTER(_IAudioEndpointVolumeCallback)


class IAudioEndpointVolumeCallback(comtypes.COMObject):
    _com_interfaces_ = [_IAudioEndpointVolumeCallback]

    def __init__(self, endpoint):
        self.__endpoint = endpoint
        comtypes.COMObject.__init__(self)

    def OnNotify(self, pNotify):
        pNotify = pNotify.contents
        ON_ENDPOINT_VOLUME_CHANGED.signal(
            device=self.__endpoint.device,
            endpoint=self.__endpoint,
            is_muted=pNotify.is_muted,
            master_volume=pNotify.master_volume,
            channel_volumes=list(pNotify)
        )
        return S_OK


class EndpointChannelVolume(object):

    def __init__(self, endpoint_volume, chan_num):
        self.__endpoint_volume = endpoint_volume
        self.__channel_number = chan_num

    @property
    def peak(self):
        endpoint_volume = self.__endpoint_volume

        # noinspection PyProtectedMember
        if endpoint_volume._peak_meter is None:
            return -1.0

        # noinspection PyProtectedMember
        peaks = list(endpoint_volume._peak_meter)
        if peaks:
            return peaks[self.channel_number]
        return -1.0

    @property
    def endpoint(self):
        endpoint_volume = self.__endpoint_volume

        return endpoint_volume.endpoint

    @property
    def channel_number(self):
        return self.__channel_number

    @property
    def level_db(self):
        endpoint_volume = self.__endpoint_volume

        pfLevel = FLOAT()
        endpoint_volume.GetChannelVolumeLevel(UINT(self.channel_number), ctypes.byref(pfLevel))
        return pfLevel.value

    @level_db.setter
    def level_db(self, value):
        endpoint_volume = self.__endpoint_volume

        endpoint_volume.SetChannelVolumeLevel(UINT(self.channel_number), FLOAT(value), NULL)

    @property
    def level(self):
        endpoint_volume = self.__endpoint_volume

        pfLevel = FLOAT()
        endpoint_volume.GetChannelVolumeLevelScalar(UINT(self.channel_number), ctypes.byref(pfLevel))
        return pfLevel.value

    @level.setter
    def level(self, value):
        endpoint_volume = self.__endpoint_volume
        endpoint_volume.SetChannelVolumeLevelScalar(UINT(self.channel_number), FLOAT(value), NULL)

    @property
    def db_minimum(self):
        endpoint_volume = self.__endpoint_volume

        if isinstance(endpoint_volume, IAudioEndpointVolumeEx):
            pflVolumeMindB = FLOAT()
            pflVolumeMaxdB = FLOAT()
            pflVolumeIncrementdB = FLOAT()

            # noinspection PyUnresolvedReferences
            endpoint_volume.GetVolumeRangeChannel(
                UINT(self.channel_number),
                ctypes.byref(pflVolumeMindB),
                ctypes.byref(pflVolumeMaxdB),
                ctypes.byref(pflVolumeIncrementdB),
            )
            return pflVolumeMindB.value

        return endpoint_volume.db_minimum

    @property
    def db_maximum(self):
        endpoint_volume = self.__endpoint_volume

        if isinstance(endpoint_volume, IAudioEndpointVolumeEx):
            pflVolumeMindB = FLOAT()
            pflVolumeMaxdB = FLOAT()
            pflVolumeIncrementdB = FLOAT()

            # noinspection PyUnresolvedReferences
            endpoint_volume.GetVolumeRangeChannel(
                UINT(self.channel_number),
                ctypes.byref(pflVolumeMindB),
                ctypes.byref(pflVolumeMaxdB),
                ctypes.byref(pflVolumeIncrementdB),
            )
            return pflVolumeMaxdB.value

        return endpoint_volume.db_maximum

    @property
    def db_increment(self):
        endpoint_volume = self.__endpoint_volume

        if isinstance(endpoint_volume, IAudioEndpointVolumeEx):
            pflVolumeMindB = FLOAT()
            pflVolumeMaxdB = FLOAT()
            pflVolumeIncrementdB = FLOAT()

            # noinspection PyUnresolvedReferences
            endpoint_volume.GetVolumeRangeChannel(
                UINT(self.channel_number),
                ctypes.byref(pflVolumeMindB),
                ctypes.byref(pflVolumeMaxdB),
                ctypes.byref(pflVolumeIncrementdB),
            )
            return pflVolumeIncrementdB.value

        return endpoint_volume.db_increment


class IAudioEndpointVolume(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioEndpointVolume
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'RegisterControlChangeNotify',
            (['in'], PIAudioEndpointVolumeCallback, 'pNotify')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'UnregisterControlChangeNotify',
            (['in'], PIAudioEndpointVolumeCallback, 'pNotify')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetChannelCount',
            (['out'], LPUINT, 'pnChannelCount')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetMasterVolumeLevel',
            (['in'], FLOAT, 'fLevelDB'),
            (['in'], LPCGUID, 'pguidEventContext', None)
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetMasterVolumeLevelScalar',
            (['in'], FLOAT, 'fLevel'),
            (['in'], LPCGUID, 'pguidEventContext', None)
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetMasterVolumeLevel',
            (['in'], LPFLOAT, 'pfLevelDB')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetMasterVolumeLevelScalar',
            (['out'], LPFLOAT, 'pfLevel')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetChannelVolumeLevel',
            (['in'], UINT, 'nChannel'),
            (['in'], FLOAT, 'fLevelDB'),
            (['in'], LPCGUID, 'pguidEventContext', None)
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetChannelVolumeLevelScalar',
            (['in'], UINT, 'nChannel'),
            (['in'], FLOAT, 'fLevel'),
            (['in'], LPCGUID, 'pguidEventContext', None)
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetChannelVolumeLevel',
            (['in'], UINT, 'nChannel'),
            (['in'], LPFLOAT, 'pfLevelDB')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetChannelVolumeLevelScalar',
            (['in'], UINT, 'nChannel'),
            (['in'], LPFLOAT, 'pfLevel')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetMute',
            (['in'], BOOL, 'bMute'),
            (['in'], LPCGUID, 'pguidEventContext', None)
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetMute',
            (['out', 'retval'], LPBOOL, 'pbMute')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetVolumeStepInfo',
            (['in'], LPUINT, 'pnStep'),
            (['in'], LPUINT, 'pnStepCount')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'VolumeStepUp',
            (['in'], LPCGUID, 'pguidEventContext', None)
        ),
        COMMETHOD(
            [],
            HRESULT,
            'VolumeStepDown',
            (['in'], LPCGUID, 'pguidEventContext', None)
        ),
        COMMETHOD(
            [],
            HRESULT,
            'QueryHardwareSupport',
            (['out'], LPDWORD, 'pdwHardwareSupportMask')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetVolumeRange',
            (['in'], LPFLOAT, 'pfLevelMinDB'),
            (['in'], LPFLOAT, 'pfLevelMaxDB'),
            (['in'], LPFLOAT, 'pfVolumeIncrementDB')
        )
    )

    def __init__(self):
        self.__device = None
        self.__endpoint = None
        self.__volume_callback = None
        self._peak_meter = None
        comtypes.IUnknown.__init__(self)

    def __call__(self, device=None, endpoint=None):
        self.__device = device
        self.__endpoint = endpoint

        if endpoint is not None:
            self.__volume_callback = IAudioEndpointVolumeCallback(endpoint)

            # noinspection PyUnresolvedReferences
            self.RegisterControlChangeNotify(self.__volume_callback)

            peak_meter = self.__endpoint.activate(IAudioMeterInformation)
            if peak_meter:
                self._peak_meter = peak_meter
            else:
                self._peak_meter = None

        return self

    def __del__(self):
        if self.__volume_callback is None:
            return

        try:
            # noinspection PyUnresolvedReferences
            self.UnregisterControlChangeNotify(self.__volume_callback)
        except ValueError:
            pass

    @property
    def peak(self):
        if self._peak_meter is None:
            return -1.0

        return self._peak_meter.peak_value

    def __iter__(self):
        # noinspection PyUnresolvedReferences
        count = self.GetChannelCount()
        for i in range(count):
            yield EndpointChannelVolume(self, i)

    @property
    def level_db(self):
        pfLevelDB = FLOAT()
        try:
            # noinspection PyUnresolvedReferences
            self.GetMasterVolumeLevel(ctypes.byref(pfLevelDB))
        except ValueError:
            return 0.0

        return pfLevelDB.value

    @level_db.setter
    def level_db(self, value):
        # noinspection PyUnresolvedReferences
        self.SetMasterVolumeLevel(FLOAT(value), NULL)

    @property
    def level(self):
        # noinspection PyUnresolvedReferences
        pfLevel = self.GetMasterVolumeLevelScalar()
        return pfLevel * 100.0

    @level.setter
    def level(self, value):
        # noinspection PyUnresolvedReferences
        self.SetMasterVolumeLevelScalar(FLOAT(value / 100.0), NULL)

    @property
    def mute(self):
        # noinspection PyUnresolvedReferences
        return bool(self.GetMute())

    @mute.setter
    def mute(self, value):
        # noinspection PyUnresolvedReferences
        self.SetMute(BOOL(value), NULL)

    @property
    def db_minimum(self):
        pflVolumeMindB = FLOAT()
        pflVolumeMaxdB = FLOAT()
        pflVolumeIncrementdB = FLOAT()
        # noinspection PyUnresolvedReferences
        self.GetVolumeRange(
            ctypes.byref(pflVolumeMindB),
            ctypes.byref(pflVolumeMaxdB),
            ctypes.byref(pflVolumeIncrementdB)
        )

        return pflVolumeMindB.value

    @property
    def db_maximum(self):
        pflVolumeMindB = FLOAT()
        pflVolumeMaxdB = FLOAT()
        pflVolumeIncrementdB = FLOAT()
        # noinspection PyUnresolvedReferences
        self.GetVolumeRange(
            ctypes.byref(pflVolumeMindB),
            ctypes.byref(pflVolumeMaxdB),
            ctypes.byref(pflVolumeIncrementdB)
        )

        return pflVolumeMaxdB.value

    @property
    def db_increment(self):
        pflVolumeMindB = FLOAT()
        pflVolumeMaxdB = FLOAT()
        pflVolumeIncrementdB = FLOAT()
        # noinspection PyUnresolvedReferences
        self.GetVolumeRange(
            ctypes.byref(pflVolumeMindB),
            ctypes.byref(pflVolumeMaxdB),
            ctypes.byref(pflVolumeIncrementdB)
        )

        return pflVolumeIncrementdB.value

    def step_up(self):
        # noinspection PyUnresolvedReferences
        self.VolumeStepUp(NULL)

    def step_down(self):
        # noinspection PyUnresolvedReferences
        self.VolumeStepDown(NULL)

    @property
    def current_step(self):
        pnStep = UINT()
        pnStepCount = UINT()
        # noinspection PyUnresolvedReferences
        self.GetVolumeStepInfo(ctypes.byref(pnStep), ctypes.byref(pnStepCount))
        return pnStep

    @property
    def number_of_steps(self):
        pnStep = UINT()
        pnStepCount = UINT()
        # noinspection PyUnresolvedReferences
        self.GetVolumeStepInfo(ctypes.byref(pnStep), ctypes.byref(pnStepCount))
        return pnStepCount

    @property
    def has_hardware_volume_control(self):
        # noinspection PyUnresolvedReferences
        return bool(self.QueryHardwareSupport() & ENDPOINT_HARDWARE_SUPPORT_VOLUME)

    @property
    def has_hardware_mute_control(self):
        # noinspection PyUnresolvedReferences
        return bool(self.QueryHardwareSupport() & ENDPOINT_HARDWARE_SUPPORT_MUTE)

    @property
    def has_hardware_peak_meter(self):
        # noinspection PyUnresolvedReferences
        return bool(self.QueryHardwareSupport() & ENDPOINT_HARDWARE_SUPPORT_METER)

    def __float__(self):
        return self.level

    def __int__(self):
        return int(self.level)

    def __str__(self):
        return str(self.level)


# noinspection PyTypeChecker
PIAudioEndpointVolume = POINTER(IAudioEndpointVolume)


class IAudioEndpointVolumeEx(IAudioEndpointVolume):
    _case_insensitive_ = False
    _iid_ = IID_IAudioEndpointVolumeEx
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetVolumeRangeChannel',
            (['in'], UINT, 'iChannel'),
            (['in'], LPFLOAT, 'pflVolumeMindB'),
            (['in'], LPFLOAT, 'pflVolumeMaxdB'),
            (['in'], LPFLOAT, 'pflVolumeIncrementdB')
        ),
    )


# noinspection PyTypeChecker
PIAudioEndpointVolumeEx = POINTER(IAudioEndpointVolumeEx)


class IAudioMeterInformation(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioMeterInformation
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetPeakValue',
            (['in'], LPFLOAT, 'pfPeak')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetMeteringChannelCount',
            (['out'], LPUINT, 'pnChannelCount')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetChannelsPeakValues',
            (['in'], UINT32, 'u32ChannelCount'),
            (['in'], LPVOID, 'afPeakValues')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'QueryHardwareSupport',
            (['out'], LPDWORD, 'pdwHardwareSupportMask')
        )
    )

    @property
    def is_suppported(self):
        # noinspection PyUnresolvedReferences
        return bool(self.QueryHardwareSupport() & ENDPOINT_HARDWARE_SUPPORT_METER)

    def __iter__(self):
        # noinspection PyUnresolvedReferences
        u32ChannelCount = self.GetMeteringChannelCount()
        afPeakValues = (FLOAT * u32ChannelCount)()
        # noinspection PyUnresolvedReferences
        self.GetChannelsPeakValues(UINT32(u32ChannelCount), ctypes.cast(afPeakValues, LPVOID))

        for i in range(u32ChannelCount):
            if afPeakValues[i]:
                yield afPeakValues[i] * 100.0

    @property
    def peak_value(self):
        pfPeak = FLOAT()
        # noinspection PyUnresolvedReferences
        self.GetPeakValue(ctypes.byref(pfPeak))
        return pfPeak.value * 100.0


# noinspection PyTypeChecker
PIAudioMeterInformation = POINTER(IAudioMeterInformation)
