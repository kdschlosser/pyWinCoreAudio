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

try:
    from .data_types import *
except ImportError:
    from data_types import *


class AUDCLNT_SHAREMODE(ENUM):
    AUDCLNT_SHAREMODE_SHARED = 1
    AUDCLNT_SHAREMODE_EXCLUSIVE = 2


PAUDCLNT_SHAREMODE = POINTER(AUDCLNT_SHAREMODE)


class AUDIO_STREAM_CATEGORY(ENUM):
    AudioCategory_Other = 0
    AudioCategory_ForegroundOnlyMedia = 1
    AudioCategory_BackgroundCapableMedia = 2
    AudioCategory_Communications = 3
    AudioCategory_Alerts = 4
    AudioCategory_SoundEffects = 5
    AudioCategory_GameEffects = 6
    AudioCategory_GameMedia = 7
    AudioCategory_GameChat = 8
    AudioCategory_Speech = 9
    AudioCategory_Movie = 10
    AudioCategory_Media = 11


PAUDIO_STREAM_CATEGORY = POINTER(AUDIO_STREAM_CATEGORY)


class STGM(ENUM):
    STGM_READ = 0x00000000


PSTGM = POINTER(STGM)


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


class AudioSessionState(ENUM):
    AudioSessionStateInactive = 0
    AudioSessionStateActive = 1
    AudioSessionStateExpired = 2


PAudioSessionState = POINTER(AudioSessionState)


class EndpointConnectorType(ENUM):
    eHostProcessConnector = 0
    eOffloadConnector = 1
    eLoopbackConnector = 2
    eKeywordDetectorConnector = 3


PEndpointConnectorType = POINTER(EndpointConnectorType)


class DeviceShareMode(ENUM):
    DeviceShared = 0
    DeviceExclusive = 1


PDeviceShareMode = POINTER(DeviceShareMode)


class AudioDeviceState(ENUM):
    Active = ENUM_VALUE(0x1, 'Active')
    Disabled = 0x2
    NotPresent = 0x4
    Unplugged = 0x8


PAudioDeviceState = POINTER(AudioDeviceState)


class EChannelMapping(ENUM):
    ePcxChanMap_FL_FR = ENUM_VALUE(0, 'FL, FR')
    ePcxChanMap_FC_LFE = ENUM_VALUE(1, 'FC, LFE')
    ePcxChanMap_BL_BR = ENUM_VALUE(2, 'BL, BR')
    ePcxChanMap_FLC_FRC = ENUM_VALUE(3, 'FLC, FRC')
    ePcxChanMap_SL_SR = ENUM_VALUE(4, 'SL, SR')
    ePcxChanMap_Unknown = ENUM_VALUE(5, 'Unknown')


PEChannelMapping = POINTER(EChannelMapping)


class EPcxConnectionType(ENUM):
    eConnTypeUnknown = ENUM_VALUE(0, 'Unknown')
    eConnType3Point5mm = ENUM_VALUE(1, '1/8"')
    eConnTypeQuarter = ENUM_VALUE(2, '1/4"')
    eConnTypeAtapiInternal = ENUM_VALUE(3, 'ATAPI Internal')
    eConnTypeRCA = ENUM_VALUE(4, 'RCA')
    eConnTypeOptical = ENUM_VALUE(5, 'Optical')
    eConnTypeOtherDigital = ENUM_VALUE(6, 'Other Digital')
    eConnTypeOtherAnalog = ENUM_VALUE(7, 'Other Analog')
    eConnTypeMultichannelAnalogDIN = ENUM_VALUE(8, 'Multichannel Analog DIN')
    eConnTypeXlrProfessional = ENUM_VALUE(9, 'XLR')
    eConnTypeRJ11Modem = ENUM_VALUE(10, 'RJ11')
    eConnTypeCombination = ENUM_VALUE(11, 'Combination')


PEPcxConnectionType = POINTER(EPcxConnectionType)


class EPcxGeoLocation(ENUM):
    eGeoLocRear = ENUM_VALUE(1, 'Rear-mounted panel')
    eGeoLocFront = ENUM_VALUE(2, 'Front-mounted panel')
    eGeoLocLeft = ENUM_VALUE(3, 'Left-mounted panel')
    eGeoLocRight = ENUM_VALUE(4, 'Right-mounted panel')
    eGeoLocTop = ENUM_VALUE(5, 'Top-mounted panel')
    eGeoLocBottom = ENUM_VALUE(6, 'Bottom-mounted panel')
    eGeoLocRearPanel = ENUM_VALUE(7, 'Rear-mounted panel')
    eGeoLocRearOPanel = ENUM_VALUE(7, 'Rear-mounted slide open or pull open panel')
    eGeoLocRiser = ENUM_VALUE(8, 'Riser card')
    eGeoLocInsideMobileLid = ENUM_VALUE(9, 'Inside Mobile Lid')
    eGeoLocDrivebay = ENUM_VALUE(10, 'Drive Bay')
    eGeoLocHDMI = ENUM_VALUE(11, 'HDMI')
    eGeoLocOutsideMobileLid = ENUM_VALUE(12, 'Outside Mobile Lid')
    eGeoLocATAPI = ENUM_VALUE(13, 'ATAPI')
    eGeoLocNotApplicable = ENUM_VALUE(14, '')
    eGeoLocReserved6 = ENUM_VALUE(15, 'Reserved')


PEPcxGeoLocation = POINTER(EPcxGeoLocation)


class EPcxGenLocation(ENUM):
    eGenLocPrimaryBox = ENUM_VALUE(0, 'On primary chassis')
    eGenLocInternal = ENUM_VALUE(1, 'Inside primary chassis')
    eGenLocSeparate = ENUM_VALUE(2, 'On separate chassis')
    eGenLocOther = ENUM_VALUE(3, 'Other location')


PEPcxGenLocation = POINTER(EPcxGenLocation)


class EPxcPortConnection(ENUM):
    ePortConnJack = ENUM_VALUE(0, 'Jack')
    ePortConnIntegratedDevice = ENUM_VALUE(1, 'Slot for an integrated device')
    ePortConnBothIntegratedAndJack = ENUM_VALUE(2, 'Jack and a slot for an integrated device')
    ePortConnUnknown = ENUM_VALUE(3, 'Unknown')


PEPxcPortConnection = POINTER(EPxcPortConnection)


class DisconnectReason(ENUM):
    DisconnectReasonDeviceRemoval = 0
    DisconnectReasonServerShutdown = 1
    DisconnectReasonFormatChanged = 2
    DisconnectReasonSessionLogoff = 3
    DisconnectReasonSessionDisconnected = 4
    DisconnectReasonExclusiveModeOverride = 5


PDisconnectReason = POINTER(DisconnectReason)


class AE_POSITION_FLAGS(ENUM):
    POSITION_INVALID = 0
    POSITION_DISCONTINUOUS = 1
    POSITION_CONTINUOUS = 2
    POSITION_QPC_ERROR = 4


PAE_POSITION_FLAGS = POINTER(AE_POSITION_FLAGS)


class APO_BUFFER_FLAGS(ENUM):
    BUFFER_INVALID = 0
    BUFFER_VALID = 1
    BUFFER_SILENT = 2


PAPO_BUFFER_FLAGS = POINTER(APO_BUFFER_FLAGS)


class AUDIO_CURVE_TYPE(ENUM):
    AUDIO_CURVE_TYPE_NONE = 0
    AUDIO_CURVE_TYPE_WINDOWS_FADE = 1


PAUDIO_CURVE_TYPE = POINTER(AUDIO_CURVE_TYPE)


class AudioSessionDisconnectReason(ENUM):
    DisconnectReasonDeviceRemoval = 0
    DisconnectReasonServerShutdown = 1
    DisconnectReasonFormatChanged = 2
    DisconnectReasonSessionLogoff = 3
    DisconnectReasonSessionDisconnected = 4
    DisconnectReasonExclusiveModeOverride = 5


PAudioSessionDisconnectReason = POINTER(AudioSessionDisconnectReason)


class KSJACK_SINK_CONNECTIONTYPE(ENUM):
    KSJACK_SINK_CONNECTIONTYPE_HDMI = ENUM_VALUE(0, 'HDMI')
    KSJACK_SINK_CONNECTIONTYPE_DISPLAYPORT = 1


PKSJACK_SINK_CONNECTIONTYPE = POINTER(KSJACK_SINK_CONNECTIONTYPE)


class KSRESET(ENUM):
    KSRESET_BEGIN = 0
    KSRESET_END = 1


PKSRESET = POINTER(KSRESET)


class KSSTATE(ENUM):
    KSSTATE_STOP = 0
    KSSTATE_ACQUIRE = 1
    KSSTATE_PAUSE = 2
    KSSTATE_RUN = 3


PKSSTATE = POINTER(KSSTATE)


class APO_CONNECTION_BUFFER_TYPE(ENUM):
    APO_CONNECTION_BUFFER_TYPE_ALLOCATED = 0
    APO_CONNECTION_BUFFER_TYPE_EXTERNAL = 1
    APO_CONNECTION_BUFFER_TYPE_DEPENDANT = 2


PAPO_CONNECTION_BUFFER_TYPE = POINTER(APO_CONNECTION_BUFFER_TYPE)


class APO_FLAG(ENUM):
    APO_FLAG_NONE = 0x00000000
    APO_FLAG_INPLACE = 0x00000001
    APO_FLAG_SAMPLESPERFRAME_MUST_MATCH = ENUM_VALUE(0x00000002, 'Samples Per Frame Must Match')
    APO_FLAG_FRAMESPERSECOND_MUST_MATCH = ENUM_VALUE(0x00000004, 'Frames Per Second Must Match')
    APO_FLAG_BITSPERSAMPLE_MUST_MATCH = ENUM_VALUE(0x00000008, 'Bits Per Sample Must Match')
    APO_FLAG_MIXER = 0x00000010
    APO_FLAG_DEFAULT = (
        APO_FLAG_SAMPLESPERFRAME_MUST_MATCH |
        APO_FLAG_FRAMESPERSECOND_MUST_MATCH |
        APO_FLAG_BITSPERSAMPLE_MUST_MATCH
    )


PAPO_FLAG = POINTER(APO_FLAG)


class AUDIO_FLOW_TYPE(ENUM):
    AUDIO_FLOW_PULL = 0
    AUDIO_FLOW_PUSH = 1


PAUDIO_FLOW_TYPE = POINTER(AUDIO_FLOW_TYPE)


class EAudioConstriction(ENUM):
    eAudioConstrictionOff = 0
    eAudioConstriction48_16 = 1
    eAudioConstriction44_16 = 2
    eAudioConstriction14_14 = 3
    eAudioConstrictionMute = 4


PEAudioConstriction = POINTER(EAudioConstriction)


class SpatialAudioMetadataWriterOverflowMode(ENUM):
    SpatialAudioMetadataWriterOverflow_Fail = 0
    SpatialAudioMetadataWriterOverflow_MergeWithNew = 1
    SpatialAudioMetadataWriterOverflow_MergeWithLast = 2


PSpatialAudioMetadataWriterOverflowMode = POINTER(
    SpatialAudioMetadataWriterOverflowMode
)


class SpatialAudioMetadataCopyMode(ENUM):
    SpatialAudioMetadataCopy_Overwrite = 0
    SpatialAudioMetadataCopy_Append = 1
    SpatialAudioMetadataCopy_AppendMergeWithLast = 2
    SpatialAudioMetadataCopy_AppendMergeWithFirst = 3


PSpatialAudioMetadataCopyMode = POINTER(
    SpatialAudioMetadataCopyMode
)


class AudioObjectType(ENUM):
    AudioObjectType_None = 0
    AudioObjectType_Dynamic = 1 << 0
    AudioObjectType_FrontLeft = 1 << 1
    AudioObjectType_FrontRight = 1 << 2
    AudioObjectType_FrontCenter = 1 << 3
    AudioObjectType_LowFrequency = 1 << 4
    AudioObjectType_SideLeft = 1 << 5
    AudioObjectType_SideRight = 1 << 6
    AudioObjectType_BackLeft = 1 << 7
    AudioObjectType_BackRight = 1 << 8
    AudioObjectType_TopFrontLeft = 1 << 9
    AudioObjectType_TopFrontRight = 1 << 10
    AudioObjectType_TopBackLeft = 1 << 11
    AudioObjectType_TopBackRight = 1 << 12
    AudioObjectType_BottomFrontLeft = 1 << 13
    AudioObjectType_BottomFrontRight = 1 << 14
    AudioObjectType_BottomBackLeft = 1 << 15
    AudioObjectType_BottomBackRight = 1 << 16
    AudioObjectType_BackCenter = 1 << 17


PAudioObjectType = POINTER(AudioObjectType)


class SpatialAudioHrtfDirectivityType(ENUM):
    SpatialAudioHrtfDirectivity_OmniDirectional = 0
    SpatialAudioHrtfDirectivity_Cardioid = 1
    SpatialAudioHrtfDirectivity_Cone = 2


PSpatialAudioHrtfDirectivityType = POINTER(SpatialAudioHrtfDirectivityType)


class SpatialAudioHrtfEnvironmentType(ENUM):
    SpatialAudioHrtfEnvironment_Small = 0
    SpatialAudioHrtfEnvironment_Medium = 1
    SpatialAudioHrtfEnvironment_Large = 2
    SpatialAudioHrtfEnvironment_Outdoors = 3
    SpatialAudioHrtfEnvironment_Average = 4


PSpatialAudioHrtfEnvironmentType = POINTER(SpatialAudioHrtfEnvironmentType)


class SpatialAudioHrtfDistanceDecayType(ENUM):
    SpatialAudioHrtfDistanceDecay_NaturalDecay = 0
    SpatialAudioHrtfDistanceDecay_CustomDecay = 1


PSpatialAudioHrtfDistanceDecayType = POINTER(SpatialAudioHrtfDistanceDecayType)


class KSPROPERTY_AUDIO(ENUM):

    KSPROPERTY_AUDIO_LATENCY = 1
    KSPROPERTY_AUDIO_COPY_PROTECTION = 2
    KSPROPERTY_AUDIO_CHANNEL_CONFIG = 3
    KSPROPERTY_AUDIO_VOLUMELEVEL = ENUM_VALUE(4, 'Volume Level')
    KSPROPERTY_AUDIO_POSITION = 5
    KSPROPERTY_AUDIO_DYNAMIC_RANGE = 6
    KSPROPERTY_AUDIO_QUALITY = 7
    KSPROPERTY_AUDIO_SAMPLING_RATE = 8
    KSPROPERTY_AUDIO_DYNAMIC_SAMPLING_RATE = 9
    KSPROPERTY_AUDIO_MIX_LEVEL_TABLE = 10
    KSPROPERTY_AUDIO_MIX_LEVEL_CAPS = 11
    KSPROPERTY_AUDIO_MUX_SOURCE = 12
    KSPROPERTY_AUDIO_MUTE = 13
    KSPROPERTY_AUDIO_BASS = 14
    KSPROPERTY_AUDIO_MID = 15
    KSPROPERTY_AUDIO_TREBLE = 16
    KSPROPERTY_AUDIO_BASS_BOOST = 17
    KSPROPERTY_AUDIO_EQ_LEVEL = 18
    KSPROPERTY_AUDIO_NUM_EQ_BANDS = 19
    KSPROPERTY_AUDIO_EQ_BANDS = 20
    KSPROPERTY_AUDIO_AGC = 21
    KSPROPERTY_AUDIO_DELAY = 22
    KSPROPERTY_AUDIO_LOUDNESS = 23
    KSPROPERTY_AUDIO_WIDE_MODE = 24
    KSPROPERTY_AUDIO_WIDENESS = 25
    KSPROPERTY_AUDIO_REVERB_LEVEL = 26
    KSPROPERTY_AUDIO_CHORUS_LEVEL = 27
    KSPROPERTY_AUDIO_DEV_SPECIFIC = 28
    KSPROPERTY_AUDIO_DEMUX_DEST = 29
    KSPROPERTY_AUDIO_STEREO_ENHANCE = 30
    KSPROPERTY_AUDIO_MANUFACTURE_GUID = 32
    KSPROPERTY_AUDIO_PRODUCT_GUID = 32
    KSPROPERTY_AUDIO_CPU_RESOURCES = ENUM_VALUE(33, 'CPU Resources')
    KSPROPERTY_AUDIO_STEREO_SPEAKER_GEOMETRY = 34
    KSPROPERTY_AUDIO_SURROUND_ENCODE = 35
    KSPROPERTY_AUDIO_3D_INTERFACE = 36
    KSPROPERTY_AUDIO_PEAKMETER = ENUM_VALUE(37, 'Peak Meter')
    KSPROPERTY_AUDIO_ALGORITHM_INSTANCE = 38
    KSPROPERTY_AUDIO_FILTER_STATE = 39
    KSPROPERTY_AUDIO_PREFERRED_STATUS = 40
    KSPROPERTY_AUDIO_PEQ_MAX_BANDS = 41
    KSPROPERTY_AUDIO_PEQ_NUM_BANDS = 42
    KSPROPERTY_AUDIO_PEQ_BAND_CENTER_FREQ = 43
    KSPROPERTY_AUDIO_PEQ_BAND_Q_FACTOR = 44
    KSPROPERTY_AUDIO_PEQ_BAND_LEVEL = 45
    KSPROPERTY_AUDIO_CHORUS_MODULATION_RATE = 46
    KSPROPERTY_AUDIO_CHORUS_MODULATION_DEPTH = 47
    KSPROPERTY_AUDIO_REVERB_TIME = 48
    KSPROPERTY_AUDIO_REVERB_DELAY_FEEDBACK = 49
    KSPROPERTY_AUDIO_POSITIONEX = ENUM_VALUE(50, 'Position EX')
    KSPROPERTY_AUDIO_MIC_ARRAY_GEOMETRY = 51
    KSPROPERTY_AUDIO_PRESENTATION_POSITION = 52
    KSPROPERTY_AUDIO_WAVERT_CURRENT_WRITE_POSITION = 53
    KSPROPERTY_AUDIO_LINEAR_BUFFER_POSITION = 54
    KSPROPERTY_AUDIO_PEAKMETER2 = ENUM_VALUE(55, 'Peak Meter 2')
    KSPROPERTY_AUDIO_WAVERT_CURRENT_WRITE_LASTBUFFER_POSITION = ENUM_VALUE(56, 'Wave RT Current Write Last Buffer Position')
    KSPROPERTY_AUDIO_VOLUMELIMIT_ENGAGED = ENUM_VALUE(57, 'Volume Limit Engaged')
    KSPROPERTY_AUDIO_MIC_SENSITIVITY = 58
    KSPROPERTY_AUDIO_MIC_SNR = ENUM_VALUE(59, 'SNR')


if __name__ == '__main__':
    for name, cls in list(globals().items())[:]:
        if type(cls) == EnumMeta:
            for key, value in cls.__dict__.items():
                if key.startswith('_'):
                    continue

                if not isinstance(value, ENUM_VALUE):
                    continue
