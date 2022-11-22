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
from .utils import remap
from .audiosessiontypes import AUDIO_STREAM_CATEGORY
from .mmdeviceapi import (
    IID_IActivateAudioInterfaceAsyncOperation,
    IID_IActivateAudioInterfaceCompletionHandler,
)


IID_IAudioCaptureClient = IID(
    '{C8ADBD64-E71E-48a0-A4DE-185C395CD317}'
)
IID_IAudioClient3 = IID(
    '{7ED4EE07-8E67-4CD4-8C1A-2B7A5987AD42}'
)
IID_IAudioClock = IID(
    '{CD63314F-3FBA-4a1b-812C-EF96358728E7}'
)
IID_IAudioClock2 = IID(
    '{6f49ff73-6727-49ac-a008-d98cf5e70048}'
)
IID_IAudioClockAdjustment = IID(
    '{f6e4c0a0-46d9-4fb8-be21-57a3ef2b626c}'
)
IID_IAudioClient = IID(
    '{1cb9ad4c-dbfa-4c32-b178-c2f568a703b2}'
)
IID_IAudioClient2 = IID(
    '{726778CD-F60A-4eda-82DE-E47610CD78AA}'
)
IID_IAudioRenderClient = IID(
    '{F294ACFC-3146-4483-A7BF-ADDCA7C260E2}'
)
IID_IAudioStreamVolume = IID(
    '{93014887-242D-4068-8A15-CF5E93B90FE3}'
)
IID_IChannelAudioVolume = IID(
    '{1C158861-B533-4B30-B1CF-E853E51C59B8}'
)
IID_ISimpleAudioVolume = IID(
    '{87CE5498-68D6-44E5-9215-6DA47EF883D8}'
)


class AUDCLNT_BUFFERFLAGS(ENUM):
    AUDCLNT_BUFFERFLAGS_DATA_DISCONTINUITY = 0x1
    AUDCLNT_BUFFERFLAGS_SILENT = 0x2
    AUDCLNT_BUFFERFLAGS_TIMESTAMP_ERROR = 0x4


PAUDCLNT_BUFFERFLAGS = POINTER(AUDCLNT_BUFFERFLAGS)


class AUDCLNT_STREAMOPTIONS(ENUM):
    AUDCLNT_STREAMOPTIONS_NONE = 0
    AUDCLNT_STREAMOPTIONS_RAW = 0x1
    AUDCLNT_STREAMOPTIONS_MATCH_FORMAT = 0x2


PAUDCLNT_STREAMOPTIONS = POINTER(AUDCLNT_STREAMOPTIONS)


class WAVEFORMATEX(ctypes.Structure):
    _fields_ = [
        ('wFormatTag', WORD),
        ('nChannels', WORD),
        ('nSamplesPerSec', WORD),
        ('nAvgBytesPerSec', WORD),
        ('nBlockAlign', WORD),
        ('wBitsPerSample', WORD),
        ('cbSize', WORD),
    ]


PWAVEFORMATEX = POINTER(WAVEFORMATEX)


class AudioClientProperties(ctypes.Structure):
    _fields_ = [
        ('cbSize', UINT32),
        ('bIsOffload', BOOL),
        ('eCategory', AUDIO_STREAM_CATEGORY),
        ('Options', AUDCLNT_STREAMOPTIONS),
    ]


PAudioClientProperties = POINTER(AudioClientProperties)


class IActivateAudioInterfaceAsyncOperation(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IActivateAudioInterfaceAsyncOperation
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetActivateResult',
            (['out'], LPHRESULT, 'activateOperation'),
            (['out'], POINTER(PIUnknown), 'activatedInterface')
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
            )
        ),
    )


# noinspection PyTypeChecker
PIActivateAudioInterfaceCompletionHandler = POINTER(
    IActivateAudioInterfaceCompletionHandler
)


class IAudioCaptureClient(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioCaptureClient
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetBuffer',
            (['out'], POINTER(LPBYTE), 'ppData'),
            (['out'], LPUINT32, 'pNumFramesToRead'),
            (['out'], LPDWORD, 'pdwFlags'),
            (['out'], LPUINT64, 'pu64DevicePosition'),
            (['out'], LPUINT64, 'pu64QPCPosition')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'ReleaseBuffer',
            (['in'], UINT32, 'NumFramesWritten'),
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetNextPacketSize',
            (['out'], LPUINT32, 'pNumFramesInNextPacket'),
        )
    )


# noinspection PyTypeChecker
PIAudioCaptureClient = POINTER(IAudioCaptureClient)


class IAudioClient(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioClient
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'Initialize',
            (['in'], DWORD, 'ShareMode'),
            (['in'], DWORD, 'StreamFlags'),
            (['in'], LONGLONG, 'hnsBufferDuration'),
            (['in'], LONGLONG, 'hnsPeriodicity'),
            (['in'], PWAVEFORMATEX, 'pFormat'),
            (['in'], LPCGUID, 'AudioSessionGuid')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetBufferSize',
            (['out', 'retval'], LPUINT32, 'pNumBufferFrames')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetStreamLatency',
            (['out', 'retval'], POINTER(LONGLONG), 'phnsLatency')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetCurrentPadding',
            (['out', 'retval'], LPUINT32, 'pNumPaddingFrames')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'IsFormatSupported',
            (['in'], DWORD, 'ShareMode'),
            (['in'], PWAVEFORMATEX, 'pFormat'),
            (['out', 'unique'], POINTER(PWAVEFORMATEX), 'ppClosestMatch')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetMixFormat',
            (['out', 'retval'], POINTER(PWAVEFORMATEX), 'ppDeviceFormat')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetDevicePeriod',
            (['out', 'retval'], POINTER(LONGLONG), 'phnsDefaultDevicePeriod'),
            (['out', 'retval'], POINTER(LONGLONG), 'phnsMinimumDevicePeriod')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetEventHandle',
            (['in'], HANDLE, 'eventHandle'),
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetService',
            (['in'], LPCGUID, 'iid'),
            (['out'], POINTER(PIUnknown), 'ppv')
        ),
        COMMETHOD([], HRESULT, 'Start'),
        COMMETHOD([], HRESULT, 'Stop'),
        COMMETHOD([], HRESULT, 'Reset'),

    )


# noinspection PyTypeChecker
PIAudioClient = POINTER(IAudioClient)


class IAudioClient2(IAudioClient):
    _case_insensitive_ = False
    _iid_ = IID_IAudioClient2
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'IsOffloadCapable',
            (['in'], AUDIO_STREAM_CATEGORY, 'Category'),
            (['out'], LPBOOL, 'pbOffloadCapable')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetClientProperties',
            (['in'], PAudioClientProperties, 'pProperties')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetBufferSizeLimits',
            (['in'], PWAVEFORMATEX, 'pFormat'),
            (['in'], BOOL, 'bEventDriven'),
            (['out'], LPREFERENCE_TIME, 'phnsMinBufferDuration'),
            (['out'], LPREFERENCE_TIME, 'phnsMaxBufferDuration')
        )
    )


# noinspection PyTypeChecker
PIAudioClient2 = POINTER(IAudioClient2)


class IAudioClient3(IAudioClient):
    _case_insensitive_ = False
    _iid_ = IID_IAudioClient3
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'InitializeSharedAudioStream',
            (['in'], DWORD, 'StreamFlags'),
            (['in'], UINT32, 'PeriodInFrames'),
            (['in'], PWAVEFORMATEX, 'pFormat'),
            (['in'], LPCGUID, 'AudioSessionGuid')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetCurrentSharedModeEnginePeriod',
            (['out'], POINTER(PWAVEFORMATEX), 'ppFormat'),
            (['out'], LPUINT32, 'pCurrentPeriodInFrames'),
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetSharedModeEnginePeriod',
            (['in'], PWAVEFORMATEX, 'pFormat'),
            (['out'], LPUINT32, 'pDefaultPeriodInFrames'),
            (['out'], LPUINT32, 'pFundamentalPeriodInFrames'),
            (['out'], LPUINT32, 'pMinPeriodInFrames'),
            (['out'], LPUINT32, 'pMaxPeriodInFrames')
        )
    )


# noinspection PyTypeChecker
PIAudioClient3 = POINTER(IAudioClient3)


class IAudioClock(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioClock
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetFrequency',
            (['out'], LPUINT64, 'pu64Frequency')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetPosition',
            (['out'], LPUINT64, 'pu64Position'),
            (['out'], LPUINT64, 'pu64QPCPosition'),
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetCharacteristics',
            (['out'], LPDWORD, 'pdwCharacteristics'),
        )
    )


# noinspection PyTypeChecker
PIAudioClock = POINTER(IAudioClock)


class IAudioClock2(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioClock2
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetDevicePosition',
            (['out'], LPUINT64, 'DevicePosition'),
            (['out'], LPUINT64, 'QPCPosition'),
        ),
    )


# noinspection PyTypeChecker
PIAudioClock2 = POINTER(IAudioClock2)


class IAudioClockAdjustment(comtypes.IUnknown):
    _iid_ = IID_IAudioClockAdjustment
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'SetSampleRate',
            (['in'], FLOAT, 'flSampleRate')
        ),
    )


# noinspection PyTypeChecker
PIAudioClockAdjustment = POINTER(IAudioClockAdjustment)


class IAudioRenderClient(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioRenderClient
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetBuffer',
            (['in'], UINT32, 'NumFramesRequested'),
            (['out'], POINTER(LPBYTE), 'PeriodInFrames')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'ReleaseBuffer',
            (['in'], UINT32, 'NumFramesWritten'),
            (['in'], DWORD, 'dwFlags'),
        )
    )


# noinspection PyTypeChecker
PIAudioRenderClient = POINTER(IAudioRenderClient)


def logger(func):

    def wrapper(*args, **kwargs):

        # print('DEBUG: ', func.__name__)
        return func(*args, **kwargs)

    return wrapper


class ISimpleAudioVolume(comtypes.IUnknown):
    """Volume control for an audio session"""
    _case_insensitive_ = False
    _iid_ = IID_ISimpleAudioVolume
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'SetMasterVolume',
            (['in'], FLOAT, 'fLevel'),
            (['in'], LPCGUID, 'EventContext')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetMasterVolume',
            (['out', 'retval'], LPFLOAT, 'pfLevel')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetMute',
            (['in'], BOOL, 'bMute'),
            (['in'], LPCGUID, 'EventContext')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetMute',
            (['out', 'retval'], LPBOOL, 'pbMute')
        )
    )

    def __str__(self):
        return str(self.level)

    def __float__(self):
        return self.level

    def __int__(self):
        return int(self.level)

    def __init__(self):
        self._session = None
        self._channel_volumes = []
        comtypes.IUnknown.__init__(self)

    @logger
    def __call__(self, session):
        self._session = session

        channel_audio_volume = session.QueryInterface(IChannelAudioVolume)
        for vol in channel_audio_volume:
            self._channel_volumes.append(vol)

        return self

    @property
    def level(self) -> float:
        """
        Get/Set the volume level for a session
        """
        # noinspection PyUnresolvedReferences
        vol = self.GetMasterVolume()

        endpoint_volume = self._session.endpoint.volume.level
        return vol * endpoint_volume

    @level.setter
    def level(self, value: float):
        # noinspection PyUnresolvedReferences

        endpoint_volume = self._session.endpoint.volume.level

        # print('endpoint volume:', endpoint_volume)

        if value > endpoint_volume:
            self._session.endpoint.volume.level = value

        new_volume = remap(value, 0.0, endpoint_volume, 0.0, 100.0)
        self.SetMasterVolume(FLOAT(new_volume / 100.0), NULL)
        # print('new session volume:', new_volume)

    @property
    def mute(self) -> bool:
        """
        Get/Set mute for the session
        """
        # noinspection PyUnresolvedReferences
        mute = self.GetMute()
        return bool(mute)

    @mute.setter
    def mute(self, value: bool):
        # noinspection PyUnresolvedReferences
        self.SetMute(BOOL(value), NULL)

    def __iter__(self):
        """
        Channel Volumes
        """
        for vol in self._channel_volumes:
            yield vol


# noinspection PyTypeChecker
PISimpleAudioVolume = POINTER(ISimpleAudioVolume)


class IAudioStreamVolume(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioStreamVolume
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetChannelCount',
            (['out'], LPUINT32, 'pdwCount')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetChannelVolume',
            (['in'], UINT32, 'dwIndex'),
            (['in'], FLOAT, 'fLevel')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetChannelVolume',
            (['in'], UINT32, 'dwIndex'),
            (['out'], LPFLOAT, 'pfLevel')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetAllVolumes',
            (['in'], UINT32, 'dwCount'),
            (['in'], LPFLOAT, 'pfVolumes')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetAllVolumes',
            (['in'], UINT32, 'dwCount'),
            (['out'], LPFLOAT, 'pfVolumes')
        ),
    )


# noinspection PyTypeChecker
PIAudioStreamVolume = POINTER(IAudioStreamVolume)


class ChannelAudioVolume(object):
    """
    Audio session channel volume
    """

    def __init__(self, channel_number, channel_audio_volume):
        self.__channel_number = channel_number
        self.__channel_audio_volume = channel_audio_volume

    @property
    def channel_number(self) -> int:
        """
        Channel number
        """
        return self.__channel_number

    @property
    def level(self) -> float:
        """
        Get/Set the channel volume for a session
        """
        return self.__channel_audio_volume.GetChannelVolume(
            self.__channel_number
        ) * 100.0

    @level.setter
    def level(self, value: float):
        self.__channel_audio_volume.SetChannelVolume(
            self.__channel_number,
            FLOAT(value / 100.0)
        )


class IChannelAudioVolume(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IChannelAudioVolume
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetChannelCount',
            (['out'], LPUINT32, 'pdwCount')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetChannelVolume',
            (['in'], UINT32, 'dwIndex'),
            (['in'], FLOAT, 'pfLevel')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetChannelVolume',
            (['in'], UINT32, 'dwIndex'),
            (['out'], LPFLOAT, 'pfLevel')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetAllVolumes',
            (['in'], UINT32, 'dwCount'),
            (['in'], LPFLOAT, 'pfVolumes')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetAllVolumes',
            (['in'], UINT32, 'dwCount'),
            (['out'], LPFLOAT, 'pfVolumes')
        ),
    )

    def __iter__(self):
        # noinspection PyUnresolvedReferences
        for i in range(self.GetChannelCount()):
            yield ChannelAudioVolume(i, self)


# noinspection PyTypeChecker
PIChannelAudioVolume = POINTER(IChannelAudioVolume)
