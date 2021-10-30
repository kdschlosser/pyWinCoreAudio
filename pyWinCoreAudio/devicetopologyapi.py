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
import comtypes
import ctypes
from .mmdeviceapi import (
    PDataFlow,
    DataFlow,
)
from .ksmedia import (
    PKSJACK_DESCRIPTION,
    PKSJACK_DESCRIPTION2,
    PKSJACK_SINK_INFORMATION,
    KSNODETYPE,
    SPEAKER_FRONT_LEFT,
    SPEAKER_FRONT_RIGHT,
    SPEAKER_FRONT_CENTER,
    SPEAKER_LOW_FREQUENCY,
    SPEAKER_BACK_LEFT,
    SPEAKER_BACK_RIGHT,
    SPEAKER_FRONT_LEFT_OF_CENTER,
    SPEAKER_FRONT_RIGHT_OF_CENTER,
    SPEAKER_BACK_CENTER,
    SPEAKER_SIDE_LEFT,
    SPEAKER_SIDE_RIGHT,
    SPEAKER_TOP_CENTER,
    SPEAKER_TOP_FRONT_LEFT,
    SPEAKER_TOP_FRONT_CENTER,
    SPEAKER_TOP_FRONT_RIGHT,
    SPEAKER_TOP_BACK_LEFT,
    SPEAKER_TOP_BACK_CENTER,
    SPEAKER_TOP_BACK_RIGHT
)
from .ks import PKSDATAFORMAT
from . import utils
from .signal import ON_PART_CHANGE
from .constant import S_OK


_CoTaskMemFree = ctypes.windll.ole32.CoTaskMemFree

IID_IPerChannelDbLevel = IID(
    '{C2F8E001-F205-4BC9-99BC-C13B1E048CCB}'
)
IID_IAudioVolumeLevel = IID(
    '{7FB7B48F-531D-44A2-BCB3-5AD5A134B3DC}'
)
IID_IAudioChannelConfig = IID(
    '{BB11C46F-EC28-493C-B88A-5DB88062CE98}'
)
IID_IAudioLoudness = IID(
    '{7D8B1437-DD53-4350-9C1B-1EE2890BD938}'
)
IID_IAudioInputSelector = IID(
    '{4F03DC02-5E6E-4653-8F72-A030C123D598}'
)
IID_IAudioOutputSelector = IID(
    '{BB515F69-94A7-429e-8B9C-271B3F11A3AB}'
)
IID_IAudioMute = IID(
    '{DF45AEEA-B74A-4B6B-AFAD-2366B6AA012E}'
)
IID_IAudioBass = IID(
    '{A2B1A1D9-4DB3-425D-A2B2-BD335CB3E2E5}'
)
IID_IAudioMidrange = IID(
    '{5E54B6D7-B44B-40D9-9A9E-E691D9CE6EDF}'
)
IID_IAudioTreble = IID(
    '{0A717812-694E-4907-B74B-BAFA5CFDCA7B}'
)
IID_IAudioAutoGainControl = IID(
    '{85401FD4-6DE4-4b9d-9869-2D6753A82F3C}'
)
IID_IAudioPeakMeter = IID(
    '{DD79923C-0599-45e0-B8B6-C8DF7DB6E796}'
)
IID_IDeviceSpecificProperty = IID(
    '{3B22BCBF-2586-4af0-8583-205D391B807C}'
)
IID_IKsFormatSupport = IID(
    '{3CB4A69D-BB6F-4D2B-95B7-452D2C155DB5}'
)
IID_IKsJackDescription = IID(
    '{4509F757-2D46-4637-8E62-CE7DB944F57B}'
)
IID_IKsJackDescription2 = IID(
    '{478F3A9B-E0C9-4827-9228-6F5505FFE76A}'
)
IID_IKsJackSinkInformation = IID(
    '{D9BD72ED-290F-4581-9FF3-61027A8FE532}'
)
IID_IPartsList = IID(
    '{6DAA848C-5EB0-45CC-AEA5-998A2CDA1FFB}'
)
IID_IPart = IID(
    '{AE2DE0E4-5BCA-4F2D-AA46-5D13F8FDB3A9}'
)
IID_IConnector = IID(
    '{9c2c4058-23f5-41de-877a-df3af236a09e}'
)
IID_ISubunit = IID(
    '{82149A85-DBA6-4487-86BB-EA8F7FEFCC71}'
)
IID_IControlInterface = IID(
    '{45d37c3f-5140-444a-ae24-400789f3cbf3}'
)
IID_IControlChangeNotify = IID(
    '{A09513ED-C709-4d21-BD7B-5F34C47F3947}'
)
IID_IDeviceTopology = IID(
    '{2A07407E-6497-4A18-9787-32F79BD0D98F}'
)
CLSID_DeviceTopology = IID(
    '{1DF639D0-5EC1-47AA-9379-828DC1AA8C59}'
)
IID_DevTopologyLib = (
    '{51B9A01D-8181-4363-B59C-E678F476DD0E}'
)

VT_EMPTY = 0  # Not specified
VT_NULL = 1  # Null
VT_I2 = 2  # a 2 byte integer
VT_I4 = 3  # a 4 byte integer
VT_R4 = 4  # a 4 byte real
VT_R8 = 5  # an 8 byte real
VT_CY = 6  # currency
VT_DATE = 7  # date
VT_BSTR = 8  # string
VT_DISPATCH = 9  # IDispatch
VT_ERROR = 10  # SCODE value
VT_BOOL = 11  # Boolean value; True is -1 and false is 0
VT_VARIANT = 12  # Variant pointer
VT_UNKNOWN = 13  # IUnknown pointer
VT_DECIMAL = 14  # 16-byte fixed-pointer value
VT_I1 = 16  # Character
VT_UI1 = 17  # Unsigned character
VT_UI2 = 18  # Unsigned short
VT_UI4 = 19  # Unsigned long
VT_I8 = 20  # 64-bit integer
VT_UI8 = 21  # 64-bit unsigned integer
VT_INT = 22  # Integer
VT_UINT = 23  # Unsigned integer
VT_VOID = 24  # C-style void
VT_HRESULT = 25  # HRESULT value
VT_PTR = 26  # Pointer type
VT_SAFEARRAY = 27  # Safe array; Use VT_ARRAY in VARIANT
VT_CARRAY = 28  # C-style array
VT_USERDEFINED = 29  # User-defined type
VT_LPSTR = 30  # Null-terminated string
VT_LPWSTR = 31  # Wide null-terminated string
VT_RECORD = 36  # User-defined type
VT_INT_PTR = 37  # Signed machine register size width
VT_UINT_PTR = 38  # Unsigned machine register size width
VT_FILETIME = 64  # FILETIME value
VT_BLOB = 65  # Length-prefixed bytes
VT_STREAM = 66  # The name of the stream follows
VT_STORAGE = 67  # The name of the storage follows
VT_STREAMED_OBJECT = 68  # The stream contains an object
VT_STORED_OBJECT = 69  # The storage contains an object
VT_BLOB_OBJECT = 70  # The blob contains an object
VT_CF = 71  # Clipboard format
VT_CLSID = 72  # Class ID
VT_VERSIONED_STREAM = 73  # Stream with a GUID version
VT_BSTR_BLOB = 4095  # Reserved
VT_VECTOR = 4096  # Simple counted array
VT_ARRAY = 8192  # SAFEARRAY pointer
VT_BYREF = 16384  # Void pointer for local use
VT_RESERVED = 32768  # Reserved
VT_ILLEGAL = 65535  # Illegal value
VT_ILLEGALMASKED = 4095  # Illegal masked value
VT_TYPEMASK = 4095  # Type mask


class PartType(ENUM):
    Connector = 0
    Subunit = 1


PPartType = POINTER(PartType)


class ConnectorType(ENUM):
    Unknown_Connector = ENUM_VALUE(0, 'Unknown')
    Physical_Internal = 1
    Physical_External = 2
    Software_IO = ENUM_VALUE(3, 'Software IO')
    Software_Fixed = 4
    Network = 5


PConnectorType = POINTER(ConnectorType)


class InterfaceBase(comtypes.IUnknown):
    _case_insensitive_ = False
    _methods_ = ()
    _iid_ = GUID('{00000000-0000-0000-0000-000000000000}')

    def __init__(self):
        self.__name = None
        comtypes.IUnknown.__init__(self)

    def __call__(self, name):
        self.__name = name
        return self

    @property
    def name(self):
        return self.__name


class IAudioAutoGainControl(InterfaceBase):
    _case_insensitive_ = False
    _iid_ = IID_IAudioAutoGainControl
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetEnabled',
            (['out'], LPBOOL, 'pbEnabled')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetEnabled',
            (['in'], BOOL, 'bEnable'),
            (['in', 'unique'], LPCGUID, 'pguidEventContext')
        )
    )

    @property
    def enabled(self):
        # noinspection PyUnresolvedReferences
        return bool(self.GetEnabled())

    @enabled.setter
    def enabled(self, value):
        # noinspection PyUnresolvedReferences
        self.SetEnabled(BOOL(value), NULL)


# noinspection PyTypeChecker
PIAudioAutoGainControl = POINTER(IAudioAutoGainControl)


class ChannelDbLevel(float):
    _channel_number = 0
    _minimum = 0.0
    _maximum = 0.0
    _increment = 0.0

    @property
    def channel_number(self):
        return self._channel_number

    @property
    def minimum(self):
        return self._minimum

    @property
    def maximum(self):
        return self._maximum

    @property
    def increment(self):
        return self._increment


class IPerChannelDbLevel(InterfaceBase):
    _iid_ = IID_IPerChannelDbLevel
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetChannelCount',
            (['out'], LPUINT, 'pcChannels')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetLevelRange',
            (['in'], UINT, 'nChannel'),
            (['in'], LPFLOAT, 'pfMinLevelDB'),
            (['in'], LPFLOAT, 'pfMaxLevelDB'),
            (['in'], LPFLOAT, 'pfStepping')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetLevel',
            (['in'], UINT, 'nChannel'),
            (['in'], LPFLOAT, 'pfLevelDB')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetLevel',
            (['in'], UINT, 'nChannel'),
            (['in'], FLOAT, 'fLevelDB'),
            (['in', 'unique'], LPCGUID, 'pguidEventContext')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetLevelUniform',
            (['in'], FLOAT, 'fLevelDB'),
            (['in', 'unique'], LPCGUID, 'pguidEventContext')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetLevelAllChannels',
            (['in'], (FLOAT * 0), 'aLevelsDB'),
            (['in'], ULONG, 'cChannels'),
            (['in', 'unique'], LPCGUID, 'pguidEventContext')
        )
    )

    def __iter__(self):
        # noinspection PyUnresolvedReferences
        count = self.GetChannelCount().contents

        for i in range(count):
            pfMinLevelDB = FLOAT()
            pfMaxLevelDB = FLOAT()
            pfStepping = FLOAT()
            pfLevelDB = FLOAT()

            # noinspection PyUnresolvedReferences
            self.GetLevel(i, ctypes.byref(pfLevelDB))
            # noinspection PyUnresolvedReferences
            self.GetLevelRange(
                i,
                ctypes.byref(pfMinLevelDB),
                ctypes.byref(pfMinLevelDB),
                ctypes.byref(pfMinLevelDB)
            )

            namespace = {
                '_channel_number': i,
                '_minimum': pfMinLevelDB.value,
                '_maximum': pfMaxLevelDB.value,
                '_increment': pfStepping.value
            }

            cls = type('ChannelDbLevel', (ChannelDbLevel,), namespace)
            yield cls(pfLevelDB.value)

    def set_uniform_level(self, value):
        # noinspection PyUnresolvedReferences
        self.SetLevelUniform(FLOAT(value), NULL)


# noinspection PyTypeChecker
PIPerChannelDbLevel = POINTER(IPerChannelDbLevel)


class IAudioBass(IPerChannelDbLevel):
    _iid_ = IID_IAudioBass


# noinspection PyTypeChecker
PIAudioBass = POINTER(IAudioBass)


class IAudioMidrange(IPerChannelDbLevel):
    _iid_ = IID_IAudioMidrange


# noinspection PyTypeChecker
PIAudioMidrange = POINTER(IAudioMidrange)


class IAudioTreble(IPerChannelDbLevel):
    _iid_ = IID_IAudioTreble


# noinspection PyTypeChecker
PIAudioTreble = POINTER(IAudioTreble)


class IAudioVolumeLevel(IPerChannelDbLevel):
    _iid_ = IID_IAudioVolumeLevel


# noinspection PyTypeChecker
PIAudioVolumeLevel = POINTER(IAudioVolumeLevel)


#
# KSNODETYPE_3D_EFFECTS
# KSNODETYPE_DAC
# KSNODETYPE_VOLUME
# KSNODETYPE_PROLOGIC_DECODER
# IPart::GetSubType

class IAudioChannelConfig(InterfaceBase):
    _case_insensitive_ = False
    _iid_ = IID_IAudioChannelConfig
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'SetChannelConfig',
            (['in'], DWORD, 'dwConfig'),
            (['in', 'unique'], LPCGUID, 'pguidEventContext')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetChannelConfig',
            (['out'], LPDWORD, 'pdwConfig')
        )
    )

    def __init__(self):
        InterfaceBase.__init__(self)
        self.front_left = False
        self.front_left_of_center = False
        self.front_center = False
        self.front_right_of_center = False
        self.front_right = False
        self.side_left = False
        self.side_right = False
        self.back_left = False
        self.back_center = False
        self.back_right = False
        self.high_center = False
        self.high_front_left = False
        self.high_front_center = False
        self.high_front_right = False
        self.high_back_left = False
        self.high_back_center = False
        self.high_back_right = False
        self.subwoofer = False

    def __call__(self, name):
        InterfaceBase.__call__(self, name)

        # noinspection PyUnresolvedReferences
        value = self.GetChannelConfig()
        self.front_left = value | SPEAKER_FRONT_LEFT == value
        self.front_left_of_center = value | SPEAKER_FRONT_LEFT_OF_CENTER == value
        self.front_center = value | SPEAKER_FRONT_CENTER == value
        self.front_right_of_center = value | SPEAKER_FRONT_RIGHT_OF_CENTER == value
        self.front_right = value | SPEAKER_FRONT_RIGHT == value
        self.side_left = value | SPEAKER_SIDE_LEFT == value
        self.side_right = value | SPEAKER_SIDE_RIGHT == value
        self.back_left = value | SPEAKER_BACK_LEFT == value
        self.back_center = value | SPEAKER_BACK_CENTER == value
        self.back_right = value | SPEAKER_BACK_RIGHT == value
        self.high_center = value | SPEAKER_TOP_CENTER == value
        self.high_front_left = value | SPEAKER_TOP_FRONT_LEFT == value
        self.high_front_center = value | SPEAKER_TOP_FRONT_CENTER == value
        self.high_front_right = value | SPEAKER_TOP_FRONT_RIGHT == value
        self.high_back_left = value | SPEAKER_TOP_BACK_LEFT == value
        self.high_back_center = value | SPEAKER_TOP_BACK_CENTER == value
        self.high_back_right = value | SPEAKER_TOP_BACK_RIGHT == value
        self.subwoofer = value | SPEAKER_LOW_FREQUENCY == value

        return self

    @property
    def num_channels(self):
        value = sum([
            self.front_left,
            self.front_left_of_center,
            self.front_center,
            self.front_right_of_center,
            self.front_right,
            self.side_left,
            self.side_right,
            self.back_left,
            self.back_center,
            self.back_right,
            self.high_center,
            self.high_front_left,
            self.high_front_center,
            self.high_front_right,
            self.high_back_left,
            self.high_back_center,
            self.high_back_right,
            self.subwoofer
        ])

        return value

    def save(self):

        value = 0

        if self.front_left:
            value |= SPEAKER_FRONT_LEFT
        if self.front_left_of_center:
            value |= SPEAKER_FRONT_LEFT_OF_CENTER
        if self.front_center:
            value |= SPEAKER_FRONT_CENTER
        if self.front_right_of_center:
            value |= SPEAKER_FRONT_RIGHT_OF_CENTER
        if self.front_right:
            value |= SPEAKER_FRONT_RIGHT
        if self.side_left:
            value |= SPEAKER_SIDE_LEFT
        if self.side_right:
            value |= SPEAKER_SIDE_RIGHT
        if self.back_left:
            value |= SPEAKER_BACK_LEFT
        if self.back_center:
            value |= SPEAKER_BACK_CENTER
        if self.back_right:
            value |= SPEAKER_BACK_RIGHT
        if self.high_center:
            value |= SPEAKER_TOP_CENTER
        if self.high_front_left:
            value |= SPEAKER_TOP_FRONT_LEFT
        if self.high_front_center:
            value |= SPEAKER_TOP_FRONT_CENTER
        if self.high_front_right:
            value |= SPEAKER_TOP_FRONT_RIGHT
        if self.high_back_left:
            value |= SPEAKER_TOP_BACK_LEFT
        if self.high_back_center:
            value |= SPEAKER_TOP_BACK_CENTER
        if self.high_back_right:
            value |= SPEAKER_TOP_BACK_RIGHT
        if self.subwoofer:
            value |= SPEAKER_LOW_FREQUENCY

        # noinspection PyUnresolvedReferences
        self.SetChannelConfig(DWORD(value), NULL)

    def __str__(self):
        eye_level = sum([
            self.front_left,
            self.front_left_of_center,
            self.front_center,
            self.front_right_of_center,
            self.front_right,
            self.side_left,
            self.side_right,
            self.back_left,
            self.back_center,
            self.back_right
        ])

        three_d = sum([
            self.high_center,
            self.high_front_left,
            self.high_front_center,
            self.high_front_right,
            self.high_back_left,
            self.high_back_center,
            self.high_back_right
        ])

        sw = int(self.subwoofer)
        if three_d:
            return '{0}.{1}.{2}'.format(eye_level, sw, three_d)
        if eye_level:
            return '{0}.{1}'.format(eye_level, sw)
        if sw:
            return '{0}.{1}'.format(eye_level, sw)
        return '0'


# noinspection PyTypeChecker
PIAudioChannelConfig = POINTER(IAudioChannelConfig)


class IAudioInputSelector(InterfaceBase):
    _case_insensitive_ = False
    _iid_ = IID_IAudioInputSelector
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetSelection',
            (['out'], LPUINT, 'pnIdSelected')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetSelection',
            (['in'], UINT, 'nIdSelect'),
            (['in', 'unique'], LPCGUID, 'pguidEventContext')
        )
    )

    @property
    def value(self):
        # noinspection PyUnresolvedReferences
        return self.GetSelection()

    @value.setter
    def value(self, val):
        # noinspection PyUnresolvedReferences
        self.SetSelection(UINT(val), NULL)


# noinspection PyTypeChecker
PIAudioInputSelector = POINTER(IAudioInputSelector)


class IAudioLoudness(InterfaceBase):
    _iid_ = IID_IAudioLoudness
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetEnabled',
            (['out'], LPBOOL, 'pbEnabled')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetEnabled',
            (['in'], BOOL, 'bEnabled'),
            (['in', 'unique'], LPCGUID, 'pguidEventContext')
        )
    )

    @property
    def enabled(self):
        # noinspection PyUnresolvedReferences
        return bool(self.GetEnabled())

    @enabled.setter
    def enabled(self, val):
        # noinspection PyUnresolvedReferences
        self.SetEnabled(BOOL(val), NULL)


# noinspection PyTypeChecker
PIAudioLoudness = POINTER(IAudioLoudness)


class IAudioMute(InterfaceBase):
    _case_insensitive_ = False
    _iid_ = IID_IAudioMute
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetMute',
            (['out'], LPBOOL, 'pbMuted')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetMute',
            (['in'], BOOL, 'bMuted'),
            (['in', 'unique'], LPCGUID, 'pguidEventContext')
        )
    )

    @property
    def enabled(self):
        # noinspection PyUnresolvedReferences
        return bool(self.GetMute())

    @enabled.setter
    def enabled(self, val):
        # noinspection PyUnresolvedReferences
        self.SetMute(BOOL(val), NULL)


# noinspection PyTypeChecker
PIAudioMute = POINTER(IAudioMute)


class IAudioOutputSelector(InterfaceBase):
    _case_insensitive_ = False
    _iid_ = IID_IAudioOutputSelector
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetSelection',
            (['out'], LPUINT, 'pnIdSelected')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetSelection',
            (['in'], UINT, 'nIdSelect'),
            (['in', 'unique'], LPCGUID, 'pguidEventContext')
        )
    )

    @property
    def value(self):
        # noinspection PyUnresolvedReferences
        return self.GetSelection()

    @value.setter
    def value(self, val):
        # noinspection PyUnresolvedReferences
        self.SetSelection(UINT(val), NULL)


# noinspection PyTypeChecker
PIAudioOutputSelector = POINTER(IAudioOutputSelector)


class IAudioPeakMeter(InterfaceBase):
    _case_insensitive_ = False
    _iid_ = IID_IAudioPeakMeter
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetChannelCount',
            (['out'], LPUINT, 'pcChannels')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetLevel',
            (['in'], UINT, 'nChannel'),
            (['in'], LPFLOAT, 'pfLevel')
        )
    )

    def __iter__(self):
        # noinspection PyUnresolvedReferences
        count = self.GetChannelCount()

        for i in range(count):
            pfLevel = FLOAT()
            # noinspection PyUnresolvedReferences
            self.GetLevel(UINT(i), ctypes.byref(pfLevel))
            yield pfLevel.value


# noinspection PyTypeChecker
PIAudioPeakMeter = POINTER(IAudioPeakMeter)


class IConnector(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IConnector

    def __init__(self):
        self.__device = None
        self.__endpoint = None
        comtypes.IUnknown.__init__(self)

    def __call__(self, device=None, endpoint=None):
        self.__device = device
        self.__endpoint = endpoint
        return self

    @property
    def part(self):
        part = self.QueryInterface(IPart)
        return part(device=self.__device, endpoint=self.__endpoint)

    @property
    def connected_to_connector_id(self):
        # noinspection PyUnresolvedReferences
        data = self.GetConnectorIdConnectedTo()
        connector_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return connector_id

    @property
    def connected_to_device_id(self):
        # noinspection PyUnresolvedReferences
        data = self.GetDeviceIdConnectedTo()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return device_id

    @property
    def type(self):
        # noinspection PyUnresolvedReferences
        type_ = self.GetType().value
        return ConnectorType.get(type_)

    @property
    def data_flow(self):
        # noinspection PyUnresolvedReferences
        data_flow = self.GetDataFlow().value
        return DataFlow.get(data_flow)

    def connect_to(self, connector):
        # noinspection PyUnresolvedReferences
        self.ConnectTo(ctypes.byref(connector))

    def disconnect(self):
        # noinspection PyUnresolvedReferences
        self.Disconnect()

    @property
    def is_connected(self):
        # noinspection PyUnresolvedReferences
        return bool(self.IsConnected())

    @property
    def connected_to(self):
        # noinspection PyTypeChecker
        ppConTo = POINTER(IConnector)()
        # noinspection PyUnresolvedReferences
        self.GetConnectedTo(ctypes.byref(ppConTo))

        return ppConTo(device=self.__device, endpoint=self.__endpoint)


# noinspection PyTypeChecker
PIConnector = POINTER(IConnector)


IConnector._methods_ = (
    COMMETHOD(
        [],
        HRESULT,
        'GetType',
        (['out'], PConnectorType, 'pType')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'GetDataFlow',
        (['out'], PDataFlow, 'pFlow')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'ConnectTo',
        (['in'], PIConnector, 'pConnectTo')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'Disconnect'
    ),
    COMMETHOD(
        [],
        HRESULT,
        'IsConnected',
        (['out'], LPBOOL, 'pbConnected')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'GetConnectedTo',
        (['in'], POINTER(PIConnector), 'ppConTo')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'GetConnectorIdConnectedTo',
        (['out'], POINTER(LPWSTR), 'ppwstrConnectorId')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'GetDeviceIdConnectedTo',
        (['out'], POINTER(LPWSTR), 'ppwstrDeviceId')
    )
)


class IControlInterface(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IControlInterface
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetName',
            (['out'], POINTER(LPWSTR), 'ppwstrName')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetIID',
            (['in'], LPCGUID, 'pIID')
        )
    )

    @property
    def name(self):
        # noinspection PyUnresolvedReferences
        data = self.GetName()
        name = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return name

    @property
    def iid(self):
        guid = GUID()
        # noinspection PyUnresolvedReferences
        self.GetIID(ctypes.byref(guid))
        return guid


# noinspection PyTypeChecker
PIControlInterface = POINTER(IControlInterface)


class IntWrapper(int):
    _minimum = 0
    _maximum = 0
    _increment = 0

    @property
    def minimum(self):
        return self._minimum

    @property
    def maximum(self):
        return self._maximum

    @property
    def icrement(self):
        return self._increment


class FloatWrapper(int):
    _minimum = 0.0
    _maximum = 0.0
    _increment = 0.0

    @property
    def minimum(self):
        return self._minimum

    @property
    def maximum(self):
        return self._maximum

    @property
    def icrement(self):
        return self._increment


class IDeviceSpecificProperty(InterfaceBase):
    _case_insensitive_ = False
    _iid_ = IID_IDeviceSpecificProperty
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetDataType',
            (['out'], LPVARTYPE, 'pVType')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetValue',
            (['in'], LPVOID, 'pvValue'),
            (['in', 'out'], LPDWORD, 'pcbValue')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetValue',
            (['in'], LPVOID, 'pvValue'),
            (['in'], DWORD, 'cbValue'),
            (['in', 'unique'], LPCGUID, 'pguidEventContext')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'Get4BRange',
            (['in'], LPLONG, 'plMin'),
            (['in'], LPLONG, 'plMax'),
            (['in'], LPLONG, 'plStepping')
        )
    )

    @property
    def value(self):
        # noinspection PyUnresolvedReferences
        data_type = self._GetDataType()

        mapping = {
            VT_BOOL: VARIANT_BOOL,
            VT_I1: BYTE,
            VT_I2: SHORT,
            VT_I4: LONG,
            VT_I8: LONGLONG,
            VT_INT: INT,
            VT_UI1: UBYTE,
            VT_UI2: USHORT,
            VT_UI4: ULONG,
            VT_UI8: ULONGLONG,
            VT_UINT: UINT,
            VT_R4: FLOAT,
            VT_R8: DOUBLE,
            VT_CY: LONGLONG,
            VT_DECIMAL: DOUBLE,
            VT_BSTR: BSTR,
            VT_LPSTR: LPSTR,
            VT_LPWSTR: LPWSTR
        }

        if data_type not in mapping:
            return None

        container = mapping[data_type]
        if data_type in (VT_LPWSTR, VT_LPSTR, VT_BSTR):
            pvValue = NULL
            pcbValue = DWORD()
            # noinspection PyUnresolvedReferences
            self.GetValue(pvValue, ctypes.byref(pcbValue))

            pvValue = (container * pcbValue.value)
            # noinspection PyUnresolvedReferences
            self.GetValue(pvValue, ctypes.byref(pcbValue))

            return utils.convert_to_string(pvValue)

        pvValue = container
        pcbValue = DWORD(ctypes.sizeof(pvValue))
        # noinspection PyUnresolvedReferences
        self.GetValue(ctypes.byref(pvValue), ctypes.byref(pcbValue))

        try:
            pvValue = pvValue.value
        except AttributeError:
            pass

        if data_type == VT_BOOL:
            return bool(pvValue)

        plMin = LONG()
        plMax = LONG()
        plStepping = LONG()

        # noinspection PyUnresolvedReferences
        self.Get4BRange(ctypes.byref(plMin), ctypes.byref(plMax), ctypes.byref(plStepping))

        namespace = {
            '_minimum': plMin.value,
            '_maximum': plMax.value,
            '_increment': plStepping.value
        }
        if isinstance(pvValue, float):
            cls = type('FloatWrapper', (FloatWrapper,), namespace)
        else:
            cls = type('IntWrapper', (IntWrapper,), namespace)

        return cls(pvValue)


# noinspection PyTypeChecker
PIDeviceSpecificProperty = POINTER(IDeviceSpecificProperty)


class ISubunit(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_ISubunit

    def __init__(self):
        self.__device = None
        self.__endpoint = None
        comtypes.IUnknown.__init__(self)

    def __call__(self, device=None, endpoint=None):
        self.__device = device
        self.__endpoint = endpoint
        return self

    @property
    def part(self):
        part = self.QueryInterface(IPart)
        return part(device=self.__device, endpoint=self.__endpoint)


# noinspection PyTypeChecker
PISubunit = POINTER(ISubunit)


class IPartsList(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IPartsList

    def __init__(self):
        self.__device = None
        self.__endpoint = None
        comtypes.IUnknown.__init__(self)

    def __call__(self, device=None, endpoint=None):
        self.__device = device
        self.__endpoint = endpoint
        return self

    def __iter__(self):
        # noinspection PyUnresolvedReferences
        count = self.GetCount()

        for i in range(count):
            # noinspection PyTypeChecker
            part = POINTER(IPart)()
            # noinspection PyUnresolvedReferences
            self.GetPart(i, ctypes.byref(part))
            yield part(device=self.__device, endpoint=self.__endpoint)


# noinspection PyTypeChecker
PIPartsList = POINTER(IPartsList)


class IPart(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IPart

    def __init__(self):
        self.__device = None
        self.__endpoint = None
        self.__interface_callbacks = []
        comtypes.IUnknown.__init__(self)

    @property
    def control_interface(self):

        # noinspection PyTypeChecker
        control_interface = POINTER(IControlInterface)()
        # noinspection PyUnresolvedReferences
        return self.GetControlInterface(0, ctypes.byref(control_interface))

    @property
    def connector(self):
        connector = self.QueryInterface(IConnector)
        if connector:
            return connector(device=self.__device, endpoint=self.__endpoint)

    @property
    def device_topology(self):
        # noinspection PyTypeChecker
        device_topology = POINTER(IDeviceTopology)()
        # noinspection PyUnresolvedReferences
        self.GetTopologyObject(ctypes.byref(device_topology))
        return device_topology(device=self.__device, endpoint=self.__endpoint)

    @property
    def name(self):
        # noinspection PyUnresolvedReferences
        data = self.GetName()
        name = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return name

    @property
    def global_id(self):
        # noinspection PyUnresolvedReferences
        data = self.GetGlobalId()
        g_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return g_id

    @property
    def local_id(self):
        # noinspection PyUnresolvedReferences
        return self.GetLocalId()

    @property
    def type(self):
        # noinspection PyUnresolvedReferences
        return PartType.get(self.GetPartType().value)

    @property
    def sub_type(self):
        guid = GUID()
        # noinspection PyUnresolvedReferences
        self.GetSubType(ctypes.byref(guid))

        return KSNODETYPE.get(guid, GUID())

    def __iter__(self):
        # noinspection PyUnresolvedReferences
        count = self.GetControlInterfaceCount()
        mapping = {
            IID_IAudioAutoGainControl: IAudioAutoGainControl,
            IID_IAudioBass: IAudioBass,
            IID_IAudioChannelConfig: IAudioChannelConfig,
            IID_IAudioInputSelector: IAudioInputSelector,
            IID_IAudioLoudness: IAudioLoudness,
            IID_IAudioMidrange: IAudioMidrange,
            IID_IAudioMute: IAudioMute,
            IID_IAudioOutputSelector: IAudioOutputSelector,
            IID_IAudioPeakMeter: IAudioPeakMeter,
            IID_IAudioTreble: IAudioTreble,
            IID_IAudioVolumeLevel: IAudioVolumeLevel,
            IID_IDeviceSpecificProperty: IDeviceSpecificProperty,
            IID_IKsFormatSupport: IKsFormatSupport,
            IID_IKsJackDescription: IKsJackDescription
        }
        for i in range(count):
            # noinspection PyTypeChecker
            control_interface = POINTER(IControlInterface)()
            # noinspection PyUnresolvedReferences
            self.GetControlInterface(i, ctypes.byref(control_interface))

            name = control_interface.name
            iid = control_interface.iid

            if iid in mapping:
                control_interface = self.activate(mapping[iid])
                if control_interface:
                    # noinspection PyUnresolvedReferences,PyCallingNonCallable
                    yield control_interface(name)

    @property
    def incoming(self):
        # noinspection PyUnresolvedReferences
        parts_list = self.EnumPartsIncoming()
        if parts_list:
            return parts_list(device=self.__device, endpoint=self.__endpoint)

    @property
    def outgoing(self):
        # noinspection PyUnresolvedReferences
        parts_list = self.EnumPartsOutgoing()
        if parts_list:
            return parts_list(device=self.__device, endpoint=self.__endpoint)

    def activate(self, cls):
        try:
            # noinspection PyUnresolvedReferences
            return ctypes.cast(
                self.Activate(
                    comtypes.CLSCTX_INPROC_SERVER,
                    cls._iid_
                ),
                POINTER(cls)
            )

        except comtypes.COMError:
            pass

    def __call__(self, device=None, endpoint=None):
        self.__device = device
        self.__endpoint = endpoint
        if endpoint is not None:
            for interface in self:
                change_notify = IControlChangeNotify(endpoint, interface)
                # noinspection PyUnresolvedReferences,PyProtectedMember
                self.RegisterControlChangeCallback(interface._iid_, change_notify)
                self.__interface_callbacks.append(change_notify)

        return self

    def __del__(self):
        while self.__interface_callbacks:
            # noinspection PyUnresolvedReferences
            self.UnregisterControlChangeCallback(self.__interface_callbacks.pop(0))


# noinspection PyTypeChecker
PIPart = POINTER(IPart)


class IDeviceTopology(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IDeviceTopology
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetConnectorCount',
            (['out'], LPUINT, 'pCount')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetConnector',
            (['in'], UINT, 'nIndex'),
            (['in'], POINTER(PIConnector), 'ppConnector')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetSubunitCount',
            (['out'], LPUINT, 'pCount')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetSubunit',
            (['in'], UINT, 'nIndex'),
            (['in'], POINTER(PISubunit), 'ppSubunit')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetPartById',
            (['in'], UINT, 'nId'),
            (['in'], POINTER(PIPart), 'ppPart')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetDeviceId',
            (['out'], POINTER(LPWSTR), 'ppwstrDeviceId')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetSignalPath',
            (['in'], PIPart, 'pIPartFrom'),
            (['in'], PIPart, 'pIPartTo'),
            (['in'], BOOL, 'bRejectMixedPaths'),
            (['out'], POINTER(PIPartsList), 'ppParts')
        ),
    )

    def __init__(self):
        self.__device = None
        self.__endpoint = None

        self.__subunits = []
        self.__parts = []
        self.__connectors = []

        comtypes.IUnknown.__init__(self)

    def __call__(self, device=None, endpoint=None):
        self.__device = device
        self.__endpoint = endpoint

        for subunit in self.subunits:
            part = subunit.part
            self.__subunits.append(subunit)
            self.__parts.append(part)

        for connector in self.connectors:
            part = connector.part
            self.__connectors.append(connector)
            self.__parts.append(part)

        return self

    @property
    def device_id(self):
        # noinspection PyUnresolvedReferences
        data = self.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return device_id.rsplit('\\', 1)[0]

    @property
    def connectors(self):
        res = []
        # noinspection PyUnresolvedReferences
        pCount = self.GetConnectorCount()

        for i in range(pCount):
            # noinspection PyTypeChecker
            connector = POINTER(IConnector)()
            # noinspection PyUnresolvedReferences
            self.GetConnector(i, ctypes.byref(connector))
            res.append(connector(device=self.__device, endpoint=self.__endpoint))

        return res

    @property
    def subunits(self):
        res = []
        # noinspection PyUnresolvedReferences
        pCount = self.GetSubunitCount()

        for i in range(pCount):
            # noinspection PyTypeChecker
            subunit = POINTER(ISubunit)()
            # noinspection PyUnresolvedReferences
            self.GetSubunit(i, ctypes.byref(subunit))
            res.append(subunit(device=self.__device, endpoint=self.__endpoint))

        return res

    def __del__(self):
        del self.__parts[:]
        del self.__subunits[:]
        del self.__connectors[:]


# noinspection PyTypeChecker
PIDeviceTopology = POINTER(IDeviceTopology)


class IKsFormatSupport(InterfaceBase):
    _case_insensitive_ = False
    _iid_ = IID_IKsFormatSupport
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'IsFormatSupported',
            (['in'], PKSDATAFORMAT, 'pKsFormat'),
            (['in'], DWORD, 'cbFormat'),
            (['out'], LPBOOL, 'pbSupported')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetDevicePreferredFormat',
            (['out'], POINTER(PKSDATAFORMAT), 'ppKsFormat')
        )
    )

    def is_format_supported(self, pKsFormat):
        # noinspection PyUnresolvedReferences
        return bool(
            self.IsFormatSupported(ctypes.byref(pKsFormat), DWORD(ctypes.sizeof(pKsFormat)))
        )

    @property
    def prefered_format(self):
        # noinspection PyUnresolvedReferences
        return self.GetDevicePreferredFormat()


# noinspection PyTypeChecker
PIKsFormatSupport = POINTER(IKsFormatSupport)


class IKsJackDescription(InterfaceBase):
    _case_insensitive_ = False
    _iid_ = IID_IKsJackDescription
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetJackCount',
            (['out'], LPUINT, 'pcJacks')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetJackDescription',
            (['in'], UINT, 'nJack'),
            (['out'], PKSJACK_DESCRIPTION, 'pDescription')
        )
    )

    def __iter__(self):
        # noinspection PyUnresolvedReferences
        count = self.GetJackCount()

        for i in range(count):
            # noinspection PyUnresolvedReferences
            yield self.GetJackDescription(i)


# noinspection PyTypeChecker
PIKsJackDescription = POINTER(IKsJackDescription)


class IKsJackDescription2(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IKsJackDescription2
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetJackCount',
            (['out'], LPUINT, 'pcJacks')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetJackDescription2',
            (['in'], UINT, 'nJack'),
            (['out'], PKSJACK_DESCRIPTION2, 'pDescription2')
        ),
    )

    def __iter__(self):
        # noinspection PyUnresolvedReferences
        count = self.GetJackCount()

        for i in range(count):
            # noinspection PyUnresolvedReferences
            yield self.GetJackDescription2(i)


# noinspection PyTypeChecker
PIKsJackDescription2 = POINTER(IKsJackDescription2)


class IKsJackSinkInformation(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IKsJackSinkInformation
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetJackSinkInformation',
            (['out'], PKSJACK_SINK_INFORMATION, 'pJackSinkInformation')
        ),
    )

    @property
    def jack_sink_information(self):
        # noinspection PyUnresolvedReferences
        return self.GetJackSinkInformation()


# noinspection PyTypeChecker
PIKsJackSinkInformation = POINTER(IKsJackSinkInformation)


class _IControlChangeNotify(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IControlChangeNotify
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'OnNotify',
            (['in'], DWORD, 'dwSenderProcessId')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetIID',
            (['in', 'unique'], LPCGUID, 'pguidEventContext')
        )
    )


# noinspection PyTypeChecker
PIControlChangeNotify = POINTER(_IControlChangeNotify)


class IControlChangeNotify(comtypes.COMObject):
    _com_interfaces_ = [_IControlChangeNotify]

    def __init__(self, endpoint, interface):
        self.__endpoint = endpoint
        self.__interface = interface
        comtypes.COMObject.__init__(self)

    def OnNotify(self, dwSenderProcessId):
        print('IControlChangeNotify.OnNotify')

        ON_PART_CHANGE.signal(
            device=self.__endpoint.device,
            endpoint=self.__endpoint,
            interface=self.__interface,
            process_id=dwSenderProcessId.value
        )
        return S_OK


IPartsList._methods_ = (
    COMMETHOD(
        [],
        HRESULT,
        'GetCount',
        (['out'], LPUINT, 'pCount')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'GetPart',
        (['in'], UINT, 'nIndex'),
        (['in'], POINTER(PIPart), 'ppPart')
    )
)


IPart._methods_ = (
    COMMETHOD(
        [],
        HRESULT,
        'GetName',
        (['out'], POINTER(LPWSTR), 'ppwstrName')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'GetLocalId',
        (['out'], LPUINT, 'pnId')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'GetGlobalId',
        (['out'], POINTER(LPWSTR), 'ppwstrGlobalId')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'GetPartType',
        (['out'], PPartType, 'pPartType')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'GetSubType',
        (['in'], LPCGUID, 'pSubType')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'GetControlInterfaceCount',
        (['out'], LPUINT, 'pCount')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'GetControlInterface',
        (['in'], UINT, 'nIndex'),
        (['in'], POINTER(PIControlInterface), 'ppInterfaceDesc')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'EnumPartsIncoming',
        (['out'], POINTER(PIPartsList), 'ppParts')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'EnumPartsOutgoing',
        (['out'], POINTER(PIPartsList), 'ppParts')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'GetTopologyObject',
        (['in'], POINTER(PIDeviceTopology), 'ppTopology')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'Activate',
        (['in'], DWORD, 'dwClsContext'),
        (['in'], REFIID, 'refiid'),
        (['out'], POINTER(LPVOID), 'ppvObject'),
    ),
    COMMETHOD(
        [],
        HRESULT,
        'RegisterControlChangeCallback',
        (['in'], REFIID, 'refiid'),
        (['in'], PIControlChangeNotify, 'pNotify'),
    ),
    COMMETHOD(
        [],
        HRESULT,
        'UnregisterControlChangeCallback',
        (['in'], PIControlChangeNotify, 'pNotify'),
    )
)


class DevTopologyLib(object):
    name = u'DevTopologyLib'
    _reg_typelib_ = (IID_DevTopologyLib, 1, 0)

    IPartsList = IPartsList
    IAudioVolumeLevel = IAudioVolumeLevel
    IAudioLoudness = IAudioLoudness
    # IAudioSpeakerMap = IAudioSpeakerMap
    IAudioInputSelector = IAudioInputSelector
    IAudioMute = IAudioMute
    IAudioBass = IAudioBass
    IAudioMidrange = IAudioMidrange
    IAudioTreble = IAudioTreble
    IAudioAutoGainControl = IAudioAutoGainControl
    IAudioOutputSelector = IAudioOutputSelector
    IAudioPeakMeter = IAudioPeakMeter
    IDeviceSpecificProperty = IDeviceSpecificProperty
    IKsFormatSupport = IKsFormatSupport


class DeviceTopology(comtypes.CoClass):
    _reg_clsid_ = CLSID_DeviceTopology
    _idlflags_ = []
    _reg_typelib_ = (IID_DevTopologyLib, 1, 0)
