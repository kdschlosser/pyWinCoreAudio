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

from .data_types import *
import comtypes
import ctypes
from .enum_constants import (
    PConnectorType,
    ConnectorType,
    PDataFlow,
    DataFlow,
    KSJACK_SINK_CONNECTIONTYPE,
    PPartType,
    PartType,
    EPcxGenLocation,
    EPcxGeoLocation,
    EPcxConnectionType,
    EPxcPortConnection
)

from .iid import (
    IID_IAudioAutoGainControl,
    IID_IAudioBass,
    IID_IAudioMidrange,
    IID_IAudioTreble,
    IID_IAudioChannelConfig,
    IID_IAudioInputSelector,
    IID_IAudioOutputSelector,
    IID_IAudioLoudness,
    IID_IAudioMute,
    IID_IAudioPeakMeter,
    IID_IAudioVolumeLevel,
    IID_IConnector,
    IID_IControlInterface,
    IID_IDeviceSpecificProperty,
    IID_IDeviceTopology,
    IID_IKsFormatSupport,
    IID_IKsJackDescription,
    IID_IKsJackDescription2,
    IID_IKsJackSinkInformation,
    IID_IPartsList,
    IID_IPart,
    IID_IPerChannelDbLevel,
    IID_ISubunit,
    IID_IControlChangeNotify,
    CLSID_DeviceTopology,
    IID_DevTopologyLib,
)
from .. import utils

from ..signal import ON_PART_CHANGE

from .constant import (
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
    SPEAKER_TOP_BACK_RIGHT,
    JACKDESC2_PRESENCE_DETECT_CAPABILITY,
    JACKDESC2_DYNAMIC_FORMAT_CHANGE_CAPABILITY,
    KSNODETYPE,
    S_OK
)

_CoTaskMemFree = ctypes.windll.ole32.CoTaskMemFree


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


class AudioSpeakers(object):

    def __init__(self, value):
        if value is None:
            value = 0

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


class KSDATAFORMAT(ctypes.Structure):
    """
    At the minimum, a data format is specified by the MajorFormat, the SubFormat, and the Specifier members.
    A family of similar data formats can share the same values for MajorFormat, SubFormat, and Specifier.
    In that case, the specific data format is distinguished by additional data that follows the
    Specifier member in memory.
    """
    _fields_ = [
        ('FormatSize', ULONG),
        ('Flags', ULONG),
        ('SampleSize', ULONG),
        ('Reserved', ULONG),
        ('MajorFormat', GUID),
        ('SubFormat', GUID),
        ('Specifier', GUID)
    ]

    def __init__(self, *args, **kwargs):

        ctypes.Structure.__init__(self, *args, **kwargs)

        self.FormatSize = ctypes.sizeof(KSDATAFORMAT)

    @property
    def flags(self):
        """
        Set flags to KSDATAFORMAT_ATTRIBUTES (0x2) to indicate that the KSDATAFORMAT is followed in
        memory by a KSMULTIPLE_ITEM of KSATTRIBUTE structures.
        """
        return self.Flags.value

    @flags.setter
    def flags(self, value):
        self.Flags = value

    @property
    def sample_size(self):
        """
        Specifies the sample size of the data, for fixed sample sizes, or zero,
        if the format has a variable sample size.
        """
        return self.SampleSize.value

    @sample_size.setter
    def sample_size(self, value):
        self.SampleSize = value

    @property
    def major_format(self):
        """
        Specifies the general format type.

        The data formats that are currently supported can be found in the
        KSDATAFORMAT_TYPE_XXX symbolic constants in the ksmedia.h header file
        that is included in the Windows Driver Kit (WDK).

        A data stream that has no particular format should use
        KSDATAFORMAT_TYPE_STREAM (defined in ks.h) as the value for its
        MajorFormat
        """
        return self.MajorFormat

    @major_format.setter
    def major_format(self, value):
        if not isinstance(value, GUID):
            value = GUID(value)

        self.MajorFormat = value

    @property
    def sub_format(self):
        """
        Specifies the subformat of a general format type.

        The data subformats that are currently supported can be found in the
        KSDATAFORMAT_SUBTYPE_XXX symbolic constants in the ksmedia.h header
        file that is included in the WDK.

        Major formats that do not support subformats should use the
        KSDATAFORMAT_SUBTYPE_NONE value for this member.
        """
        return self.SubFormat

    @sub_format.setter
    def sub_format(self, value):
        if not isinstance(value, GUID):
            value = GUID(value)

        self.SubFormat = value

    @property
    def specifier(self):
        """
        Specifies additional data format type information for a specific setting of MajorFormat and SubFormat.

        The significance of this field is determined by the major format (and subformat, if the
        major format supports subformats). For example, Specifier can represent a particular
        encoding of a subformat, or it can be used to specify what type of data structure
        follows KSDATAFORMAT in memory.

        The following specifiers (defined in ks.h) are of general use:

        KSDATAFORMAT_SPECIFIER_NONE: Stands for no specifier. Used for formats that do not support specifiers.
        KSDATAFORMAT_SPECIFIER_FILENAME: Indicates that a null-terminated Unicode string immediately
        follows the KSDATAFORMAT structure in memory.
        KSDATAFORMAT_SPECIFIER_FILEHANDLE: Indicates that a file handle immediately follows KSDATAFORMAT in memory.
        """
        return self.Specifier

    @specifier.setter
    def specifier(self, value):
        if not isinstance(value, GUID):
            value = GUID(value)

        self.Specifier = value


PKSDATAFORMAT = POINTER(KSDATAFORMAT)


class JackDescription(object):

    def __init__(self, jd1, jd2):
        self.__jd1 = jd1
        self.__jd2 = jd2

    @property
    def presence_detection(self):
        if self.__jd2 is None:
            return False

        return self.__jd2.presence_detection

    @property
    def dynamic_format_change(self):
        if self.__jd2 is None:
            return False

        return self.__jd2.dynamic_format_change

    @property
    def channel_mapping(self):
        return self.__jd1.channel_mapping

    @property
    def color(self):
        return self.__jd1.color

    @property
    def connection_type(self):
        return self.__jd1.connection_type

    @property
    def geo_location(self):
        return self.__jd1.geo_location

    @property
    def gen_location(self):
        return self.__jd1.gen_location

    @property
    def port_connection(self):
        return self.__jd1.port_connection

    @property
    def is_connected(self):
        return self.__jd1.is_connected

    @property
    def location(self):
        return self.__jd1.location


class KSJACK_DESCRIPTION(ctypes.Structure):
    _fields_ = [
        ('ChannelMapping', DWORD),
        ('Color', COLORREF),
        ('ConnectionType', EPcxConnectionType),
        ('GeoLocation', EPcxGeoLocation),
        ('GenLocation', EPcxGenLocation),
        ('PortConnection', EPxcPortConnection),
        ('IsConnected', BOOL)
    ]

    @property
    def channel_mapping(self):
        return AudioSpeakers(self.ChannelMapping)

    @property
    def color(self):
        return utils.convert_triplet_to_rgb(self.Color)

    @property
    def connection_type(self):
        return EPcxConnectionType.get(self.ConnectionType.value)

    @property
    def geo_location(self):
        return EPcxGeoLocation.get(self.GeoLocation.value)

    @property
    def gen_location(self):
        return EPcxGenLocation.get(self.GenLocation.value)

    @property
    def port_connection(self):
        return EPxcPortConnection.get(self.PortConnection.value)

    @property
    def is_connected(self):
        return bool(self.IsConnected)

    @property
    def location(self):
        if str(self.geo_location):
            return '{0}, {1}'.format(self.gen_location, self.geo_location)

        return self.gen_location


PKSJACK_DESCRIPTION = POINTER(KSJACK_DESCRIPTION)


class LUID(ctypes.Structure):
    _fields_ = [
        ('LowPart', DWORD),
        ('HighPart', LONG)
    ]

    @property
    def value(self):
        return self.HighPart << 8 | self.LowPart


PLUID = POINTER(LUID)


class tagKSJACK_SINK_INFORMATION(ctypes.Structure):
    _fields_ = [
        ('ConnType', KSJACK_SINK_CONNECTIONTYPE),
        ('ManufacturerId', WORD),
        ('ProductId', WORD),
        ('AudioLatency', WORD),
        ('HDCPCapable', BOOL),
        ('AICapable', BOOL),
        ('SinkDescriptionLength', UCHAR),
        ('SinkDescription', (WCHAR * 32)),
        ('PortId', LUID),
    ]

    @property
    def manufacturer_id(self):
        return self.ManufacturerId

    @property
    def product_id(self):
        return self.ProductId

    @property
    def audio_latency(self):
        return self.AudioLatency

    @property
    def hdcp_capable(self):
        return bool(self.HDCPCapable)

    @property
    def ai_capable(self):
        return bool(self.AICapable)

    @property
    def description(self):
        return utils.convert_to_string(self.SinkDescription)

    @property
    def port_id(self):
        return self.PortId.value

    @property
    def connection_type(self):
        return KSJACK_SINK_CONNECTIONTYPE.get(self.ConnType.value)


KSJACK_SINK_INFORMATION = tagKSJACK_SINK_INFORMATION
PKSJACK_SINK_INFORMATION = POINTER(KSJACK_SINK_INFORMATION)


class tagKSJACK_DESCRIPTION2(ctypes.Structure):
    _fields_ = [
        ('DeviceStateInfo', DWORD),
        ('JackCapabilities', DWORD)
    ]

    @property
    def presence_detection(self):
        return bool(self.JackCapabilities & JACKDESC2_PRESENCE_DETECT_CAPABILITY)

    @property
    def dynamic_format_change(self):
        return bool(self.JackCapabilities & JACKDESC2_DYNAMIC_FORMAT_CHANGE_CAPABILITY)


KSJACK_DESCRIPTION2 = tagKSJACK_DESCRIPTION2
PKSJACK_DESCRIPTION2 = POINTER(KSJACK_DESCRIPTION2)


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
        return bool(self.GetEnabled())

    @enabled.setter
    def enabled(self, value):
        self.SetEnabled(BOOL(value), NULL)


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
        count = self.GetChannelCount().contents

        for i in range(count):
            pfMinLevelDB = FLOAT()
            pfMaxLevelDB = FLOAT()
            pfStepping = FLOAT()
            pfLevelDB = FLOAT()

            self.GetLevel(i, ctypes.byref(pfLevelDB))
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
        self.SetLevelUniform(FLOAT(value), NULL)


PIPerChannelDbLevel = POINTER(IPerChannelDbLevel)


class IAudioBass(IPerChannelDbLevel):
    _iid_ = IID_IAudioBass


PIAudioBass = POINTER(IAudioBass)


class IAudioMidrange(IPerChannelDbLevel):
    _iid_ = IID_IAudioMidrange


PIAudioMidrange = POINTER(IAudioMidrange)


class IAudioTreble(IPerChannelDbLevel):
    _iid_ = IID_IAudioTreble


PIAudioTreble = POINTER(IAudioTreble)


class IAudioVolumeLevel(IPerChannelDbLevel):
    _iid_ = IID_IAudioVolumeLevel


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

    @property
    def value(self):
        return self.GetChannelConfig()

    @value.setter
    def value(self, val):
        self.SetChannelConfig(DWORD(val), NULL)


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
        return self.GetSelection()

    @value.setter
    def value(self, val):
        self.SetSelection(UINT(val), NULL)


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
        return bool(self.GetEnabled())

    @enabled.setter
    def enabled(self, val):
        self.SetEnabled(BOOL(val), NULL)


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
        return bool(self.GetMute())

    @enabled.setter
    def enabled(self, val):
        self.SetMute(BOOL(val), NULL)


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
        return self.GetSelection()

    @value.setter
    def value(self, val):
        self.SetSelection(UINT(val), NULL)


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
        count = self.GetChannelCount()

        for i in range(count):
            pfLevel = FLOAT()
            self.GetLevel(UINT(i), ctypes.byref(pfLevel))
            yield pfLevel.value


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
        data = self.GetConnectorIdConnectedTo()
        connector_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return connector_id

    @property
    def connected_to_device_id(self):
        data = self.GetDeviceIdConnectedTo()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return device_id

    @property
    def type(self):
        type_ = self.GetType().value
        return ConnectorType.get(type_)

    @property
    def data_flow(self):
        data_flow = self.GetDataFlow().value
        return DataFlow.get(data_flow)

    def connect_to(self, connector):
        self.ConnectTo(ctypes.byref(connector))

    def disconnect(self):
        self.Disconnect()

    @property
    def is_connected(self):
        return bool(self.IsConnected())

    @property
    def connected_to(self):
        ppConTo = POINTER(IConnector)()
        self.GetConnectedTo(ctypes.byref(ppConTo))

        return ppConTo(device=self.__device, endpoint=self.__endpoint)


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
        data = self.GetName()
        name = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return name

    @property
    def iid(self):
        guid = GUID()
        self.GetIID(ctypes.byref(guid))
        return guid


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
            self.GetValue(pvValue, ctypes.byref(pcbValue))

            pvValue = (container * pcbValue.value)
            self.GetValue(pvValue, ctypes.byref(pcbValue))

            return utils.convert_to_string(pvValue)

        pvValue = container
        pcbValue = DWORD(ctypes.sizeof(pvValue))
        self._GetValue(ctypes.byref(pvValue), ctypes.byref(pcbValue))

        try:
            pvValue = pvValue.value
        except:
            pass

        if data_type == VT_BOOL:
            return bool(pvValue)

        plMin = LONG()
        plMax = LONG()
        plStepping = LONG()

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
        count = self.GetCount()

        for i in range(count):
            part = POINTER(IPart)()
            self.GetPart(i, ctypes.byref(part))
            yield part(device=self.__device, endpoint=self.__endpoint)


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
    def device_topology(self):
        device_topology = POINTER(IDeviceTopology)()
        self.GetTopologyObject(ctypes.byref(device_topology))
        return device_topology(device=self.__device, endpoint=self.__endpoint)

    @property
    def name(self):
        data = self.GetName()
        name = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return name

    @property
    def global_id(self):
        data = self.GetGlobalId()
        g_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return g_id

    @property
    def local_id(self):
        return self.GetLocalId()

    @property
    def part_type(self):
        return PartType.get(self.GetPartType().value)

    @property
    def sub_type(self):
        guid = GUID()
        self.GetSubType(ctypes.byref(guid))

        return KSNODETYPE.get(guid, '')

    def __iter__(self):
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
            control_interface = POINTER(IControlInterface)()
            self.GetControlInterface(i, ctypes.byref(control_interface))
            name = control_interface.name
            iid = control_interface.iid

            if iid in mapping:
                control_interface = self.activate(mapping[iid])
                if control_interface:
                    yield control_interface(name)

    @property
    def incoming(self):
        parts_list = POINTER(IPartsList)()
        self.EnumPartsIncoming(ctypes.byref(parts_list))
        return parts_list(device=self.__device, endpoint=self.__endpoint)

    @property
    def outgoing(self):
        parts_list = POINTER(IPartsList)
        self.EnumPartsOutgoing(ctypes.byref(parts_list))
        return parts_list(device=self.__device, endpoint=self.__endpoint)

    def activate(self, cls):
        try:
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
                self.RegisterControlChangeCallback(interface._iid_, change_notify)
                self.__interface_callbacks.append(change_notify)

        return self

    def __del__(self):
        while self.__interface_callbacks:
            self.UnregisterControlChangeCallback(self.__interface_callbacks.pop(0))


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
        data = self.GetDeviceId()
        device_id = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return device_id.rsplit('\\', 1)[0]

    @property
    def connectors(self):
        res = []
        pCount = self.GetConnectorCount()

        for i in range(pCount):
            connector = POINTER(IConnector)()
            self.GetConnector(i, ctypes.byref(connector))
            res.append(connector(device=self.__device, endpoint=self.__endpoint))

        return res

    @property
    def subunits(self):
        res = []
        pCount = self.GetSubunitCount()

        for i in range(pCount):
            subunit = POINTER(ISubunit)()
            self.GetSubunit(i, ctypes.byref(subunit))
            res.append(subunit(device=self.__device, endpoint=self.__endpoint))

        return res

    def __del__(self):
        del self.__parts[:]
        del self.__subunits[:]
        del self.__connectors[:]


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
        return bool(
            self.IsFormatSupported(ctypes.byref(pKsFormat), DWORD(ctypes.sizeof(pKsFormat)))
        )

    @property
    def prefered_format(self):
        return self.GetDevicePreferredFormat()


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
        count = self.GetJackCount()

        for i in range(count):
            yield self.GetJackDescription(i)


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
        count = self.GetJackCount()

        for i in range(count):
            yield self.GetJackDescription2(i)


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
        return self.GetJackSinkInformation()


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


PIControlChangeNotify = POINTER(_IControlChangeNotify)


class IControlChangeNotify(comtypes.COMObject):
    _com_interfaces_ = [_IControlChangeNotify]

    def __init__(self, endpoint, interface):
        self.__endpoint = endpoint
        self.__interface = interface
        comtypes.COMObject.__init__(self)

    def OnNotify(self, dwSenderProcessId):
        ON_PART_CHANGE.signal(
            self.__endpoint.device,
            self.__endpoint,
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
        (['in'], POINTER(PIPartsList), 'ppParts')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'EnumPartsOutgoing',
        (['in'], POINTER(PIPartsList), 'ppParts')
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
