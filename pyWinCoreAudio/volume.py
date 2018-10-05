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

import ctypes
import comtypes
from utils import run_in_thread
from singleton import Singleton
from pyWinAPI.wtypes_h import FLOAT
from pyWinAPI.winerror_h import S_OK
from pyWinAPI.endpointvolume_h import (
    ENDPOINT_HARDWARE_SUPPORT_VOLUME,
    ENDPOINT_HARDWARE_SUPPORT_MUTE,
    IID_IAudioEndpointVolumeEx,
    IAudioEndpointVolumeEx,
    IAudioMeterInformation,
    IAudioEndpointVolumeCallback,
    IID_IAudioMeterInformation
)


class AudioEndpointVolumeCallback(comtypes.COMObject):
    _com_interfaces_ = [IAudioEndpointVolumeCallback]

    def __init__(self, endpoint, callback):
        self.__endpoint = endpoint
        self.__callback = callback
        comtypes.COMObject.__init__(self)

    def OnNotify(self, pNotify):
        try:
            bMuted = bool(pNotify.contents.bMuted)
            fMasterVolume = pNotify.contents.fMasterVolume
            nChannels = pNotify.contents.nChannels
            pfChannelVolumes = ctypes.cast(
                pNotify.contents.afChannelVolumes,
                ctypes.POINTER(FLOAT)
            )
            pfChannelVolumes = list(
                pfChannelVolumes[i] for i in range(nChannels)
            )

            def do(volume, mute, channel_volumes):
                self.__callback.endpoint_volume_change(
                    self.__endpoint,
                    volume,
                    channel_volumes,
                    mute
                )

            run_in_thread(do, fMasterVolume, bMuted, pfChannelVolumes)
        except:
            import traceback
            traceback.print_exc()
            pass

        return S_OK


class AudioVolume(object):
    __metaclass__ = Singleton

    def __init__(self, endpoint):
        self.__endpoint = endpoint
        self.__volume = endpoint.activate(
            IID_IAudioEndpointVolumeEx,
            ctypes.POINTER(IAudioEndpointVolumeEx)
        )
        try:
            support = self.__volume.QueryHardwareSupport()
            if support | ENDPOINT_HARDWARE_SUPPORT_VOLUME != support:
                raise NotImplementedError
        except ValueError:
            raise NotImplementedError

    @property
    def endpoint(self):
        return self.__endpoint

    @property
    def channel_count(self):
        return self.__volume.GetChannelCount()

    @property
    def channel_levels(self):
        res = []
        for i in range(self.channel_count):
            res += [self.__volume.GetChannelVolumeLevel(i)]
        return res

    @channel_levels.setter
    def channel_levels(self, levels):
        for i, level in enumerate(levels):
            self.__volume.SetChannelVolumeLevel(i, level)

    @property
    def channel_levels_scalar(self):
        res = []
        for i in range(self.channel_count):
            res += [self.__volume.GetChannelVolumeLevelScalar(i)]
        return res

    @channel_levels_scalar.setter
    def channel_levels_scalar(self, levels):
        for i, level in enumerate(levels):
            self.__volume.SetChannelVolumeLevelScalar(i, level)

    @property
    def master(self):
        return self.__volume.GetMasterVolumeLevel()

    @master.setter
    def master(self, level):
        self.__volume.SetMasterVolumeLevel(level)

    @property
    def master_scalar(self):
        return self.__volume.GetMasterVolumeLevelScalar()

    @master_scalar.setter
    def master_scalar(self, level):
        self.__volume.SetMasterVolumeLevelScalar(level)

    @property
    def mute(self):
        support = self.__volume.QueryHardwareSupport()

        if support | ENDPOINT_HARDWARE_SUPPORT_MUTE == support:
            return bool(self.__volume.GetMute())
        raise AttributeError

    @mute.setter
    def mute(self, mute):
        support = self.__volume.QueryHardwareSupport()

        if support | ENDPOINT_HARDWARE_SUPPORT_MUTE == support:
            self.__volume.SetMute(mute)
        else:
            raise AttributeError

    @property
    def range(self):
        return self.__volume.GetVolumeRange()

    @property
    def channel_ranges(self):
        res = []
        for i in range(self.channel_count):
            res += [self.__volume.GetVolumeRangeChannel(i)]
        return res

    @property
    def step(self):
        return self.__volume.GetVolumeStepInfo()

    def register_notification_callback(self, callback):
        volume_callback = AudioEndpointVolumeCallback(
            self.__endpoint,
            callback
        )
        self.__volume.RegisterControlChangeNotify(volume_callback)
        return volume_callback

    def unregister_notification_callback(self, callback):
        self.__volume.UnregisterControlChangeNotify(callback)

    def up(self):
        self.__volume.VolumeStepUp()

    def down(self):
        self.__volume.VolumeStepDown()

    @property
    def peak_meter(self):
        return AudioPeakMeter(self.__endpoint)


class AudioPeakMeter(object):

    def __init__(self, endpoint):
        self.__peak_meter = endpoint.activate(
            IID_IAudioMeterInformation,
            ctypes.POINTER(IAudioMeterInformation)
        )

    @property
    def channel_peak_values(self):

        # support = self.__peak_meter.QueryHardwareSupport()
        # if support | ENDPOINT_HARDWARE_SUPPORT_METER != support:
            # raise NotImplementedError

        channel_count = self.channel_count
        channels = self.__peak_meter.GetChannelsPeakValues(channel_count)

        return list(channels[i] for i in range(channel_count))

    @property
    def channel_count(self):
        return self.__peak_meter.GetMeteringChannelCount()

    @property
    def peak_value(self):
        return self.__peak_meter.GetPeakValue()
