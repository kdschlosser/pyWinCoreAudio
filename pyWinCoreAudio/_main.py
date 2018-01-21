
import comtypes
from _audioclient import PIAudioClient
from _audiopolicy import PIAudioSessionManager2
from _endpointvolumeapi import (
    PIAudioEndpointVolumeEx,
    PIAudioMeterInformation
)
from _mmdeviceapi import (
    IMMDeviceEnumerator,
    IMMNotificationClient,
    IMMEndpoint,
)
from _enum import (
    EDataFlow,
    ERole,
    EndpointConnectorType,
    EndpointFormFactor,
    ConnectorType,
    PartType,
    AudioDeviceState,
    EPcxConnectionType,
    EPcxGeoLocation,
    EPcxGenLocation,
    EPxcPortConnection,
    KSJACK_SINK_CONNECTIONTYPE
)
from _devicetopologyapi import (
    PIDeviceTopology,
    IPart,
    PIAudioBass,
    PIDeviceSpecificProperty,
    PIKsJackDescription,
    PIKsJackDescription2,
    PIKsJackSinkInformation,
    PIAudioOutputSelector,
    PIAudioInputSelector,
    PIAudioChannelConfig,
    PIAudioAutoGainControl,
    PIAudioPeakMeter,
    PIAudioMute,
    PIAudioMidrange,
    PIAudioLoudness,
    PIKsFormatSupport,
    PIAudioVolumeLevel,
    PIAudioTreble,
)
from _iid import (
    IID_IAudioAutoGainControl,
    IID_IAudioBass,
    IID_IAudioChannelConfig,
    IID_IAudioInputSelector,
    IID_IAudioLoudness,
    IID_IAudioMidrange,
    IID_IAudioMute,
    IID_IAudioOutputSelector,
    IID_IAudioPeakMeter,
    IID_IAudioTreble,
    IID_IAudioVolumeLevel,
    IID_IAudioMeterInformation,
    IID_IAudioSessionManager2,
    IID_IAudioClient,
    IID_IAudioEndpointVolumeEx,
    IID_IDeviceTopology,
    IID_IDeviceSpecificProperty,
    IID_IKsFormatSupport,
    IID_IKsJackDescription,
    IID_IKsJackDescription2,
    IID_IKsJackSinkInformation,
    CLSID_MMDeviceEnumerator,
)
from _constant import (
    PKEY_Device_FriendlyName,
    PKEY_Device_DeviceDesc,
    PKEY_DeviceInterface_FriendlyName,
    PKEY_AudioEndpoint_Disable_SysFx,
    PKEY_AudioEndpoint_PhysicalSpeakers,
    PKEY_AudioEndpoint_GUID,
    PKEY_AudioEndpoint_FullRangeSpeakers,
    PKEY_AudioEndpoint_FormFactor,
    PKEY_AudioEndpoint_JackSubType,
    STGM_READ,
    DEVICE_STATE_MASK_ALL,
    ENDPOINT_HARDWARE_SUPPORT_VOLUME,
    ENDPOINT_HARDWARE_SUPPORT_MUTE,
    ENDPOINT_HARDWARE_SUPPORT_METER,
    KSAUDIO_SPEAKER_DIRECTOUT,
    KSAUDIO_SPEAKER_MONO,
    KSAUDIO_SPEAKER_1POINT1,
    KSAUDIO_SPEAKER_STEREO,
    KSAUDIO_SPEAKER_2POINT1,
    KSAUDIO_SPEAKER_3POINT0,
    KSAUDIO_SPEAKER_3POINT1,
    KSAUDIO_SPEAKER_QUAD,
    KSAUDIO_SPEAKER_SURROUND,
    KSAUDIO_SPEAKER_5POINT0,
    KSAUDIO_SPEAKER_5POINT1,
    KSAUDIO_SPEAKER_7POINT0,
    KSAUDIO_SPEAKER_7POINT1,
    KSAUDIO_SPEAKER_5POINT1_SURROUND,
    KSAUDIO_SPEAKER_7POINT1_SURROUND,
)

CLSCTX_INPROC_SERVER = comtypes.CLSCTX_INPROC_SERVER

SINK_CONNECTIONTYPE = {
    KSJACK_SINK_CONNECTIONTYPE.KSJACK_SINK_CONNECTIONTYPE_HDMI:        'HDMI',
    KSJACK_SINK_CONNECTIONTYPE.KSJACK_SINK_CONNECTIONTYPE_DISPLAYPORT: (
        'Displayport'
    )
}

ICONTROLS = {
    IID_IAudioAutoGainControl: PIAudioAutoGainControl,
    IID_IAudioBass: PIAudioBass,
    IID_IAudioChannelConfig: PIAudioChannelConfig,
    IID_IAudioInputSelector: PIAudioInputSelector,
    IID_IAudioLoudness: PIAudioLoudness,
    IID_IAudioMidrange: PIAudioMidrange,
    IID_IAudioMute: PIAudioMute,
    IID_IAudioOutputSelector: PIAudioOutputSelector,
    IID_IAudioPeakMeter: PIAudioPeakMeter,
    IID_IAudioTreble: PIAudioTreble,
    IID_IAudioVolumeLevel: PIAudioVolumeLevel,
    IID_IAudioMeterInformation: PIAudioMeterInformation,
    IID_IKsFormatSupport: PIKsFormatSupport,
    IID_IKsJackDescription: PIKsJackDescription,
    IID_IKsJackDescription2: PIKsJackDescription2,
    IID_IDeviceSpecificProperty: PIDeviceSpecificProperty,
    IID_IKsJackSinkInformation: PIKsJackSinkInformation
}

ENDPOINT_CONNECTOR_TYPE = {
    EndpointConnectorType.eHostProcessConnector:     'Host Process Connector',
    EndpointConnectorType.eOffloadConnector:         'Offload Connector',
    EndpointConnectorType.eLoopbackConnector:        'Loopback Connector',
    EndpointConnectorType.eConnectorCount:           'Connector Count',
    EndpointConnectorType.eKeywordDetectorConnector: (
        'Keyword Detector Connector'
    ),
}

ENDPOINT_FORM_FACTOR = {
    EndpointFormFactor.RemoteNetworkDevice:           'Remote Network Device',
    EndpointFormFactor.Speakers:                      'Speakers',
    EndpointFormFactor.LineLevel:                     'Line Level',
    EndpointFormFactor.Headphones:                    'Headphones',
    EndpointFormFactor.Microphone:                    'Microphone',
    EndpointFormFactor.Headset:                       'Headset',
    EndpointFormFactor.Handset:                       'Handset',
    EndpointFormFactor.SPDIF:                         'SPDIF',
    EndpointFormFactor.UnknownFormFactor:             'Unknown Form Factor',
    EndpointFormFactor.DigitalAudioDisplayDevice:     (
        'Digital Audio Display Device'
    ),
    EndpointFormFactor.UnknownDigitalPassthrough:     (
        'Unknown Digital Passthrough'
    ),
    EndpointFormFactor.EndpointFormFactor_enum_count: (
        'EndpointFormFactor_enum_count'
    )
}

CONNECTOR_TYPE = {
    ConnectorType.Unknown_Connector: 'Unknown Connector',
    ConnectorType.Physical_Internal: 'Physical Internal',
    ConnectorType.Physical_External: 'Physical External',
    ConnectorType.Software_IO:       'Software IO',
    ConnectorType.Software_Fixed:    'Software Fixed',
    ConnectorType.Network:           'Network'
}

AUDIO_DEVICE_STATE = {
    AudioDeviceState.Active:     'Active',
    AudioDeviceState.Disabled:   'Disabled',
    AudioDeviceState.NotPresent: 'Not Present',
    AudioDeviceState.Unplugged:  'Unplugged'
}

E_DATA_FLOW = {
    EDataFlow.eRender:              'Render',
    EDataFlow.eCapture:             'Capture',
    EDataFlow.eAll:                 'All',
    EDataFlow.EDataFlow_enum_count: 'EDataFlow_enum_count'
}

E_ROLE = {
    ERole.eConsole:         'Console',
    ERole.eMultimedia:      'Multimedia',
    ERole.eCommunications:  'Communications',
    ERole.ERole_enum_count: 'ERole_enum_count'
}


def GetRGB(triplet):
    if not triplet:
        return 0, 0, 0

    r, g, b = bytearray.fromhex(hex(triplet)[2:].replace('L', '').zfill(6))
    return r, g, b


KSAUDIO_SPEAKER = {
    KSAUDIO_SPEAKER_DIRECTOUT: dict(
        front_left=False,
        front_right=False,
        center=False,
        high_left=False,
        high_right=False,
        side_left=False,
        side_right=False,
        back_left=False,
        back_right=False,
        back_center=False,
        subwoofer=False,
        string='Direct'
    ),
    KSAUDIO_SPEAKER_MONO: dict(
        front_left=False,
        front_right=False,
        center=True,
        high_left=False,
        high_right=False,
        side_left=False,
        side_right=False,
        back_left=False,
        back_right=False,
        back_center=False,
        subwoofer=False,
        string='Mono'
    ),
    KSAUDIO_SPEAKER_1POINT1: dict(
        front_left=False,
        front_right=False,
        center=True,
        high_left=False,
        high_right=False,
        side_left=False,
        side_right=False,
        back_left=False,
        back_right=False,
        back_center=False,
        subwoofer=True,
        string='Mono with subwoofer'
    ),
    KSAUDIO_SPEAKER_STEREO: dict(
        front_left=True,
        front_right=True,
        center=False,
        high_left=False,
        high_right=False,
        side_left=False,
        side_right=False,
        back_left=False,
        back_right=False,
        back_center=False,
        subwoofer=False,
        string='Stereo'
    ),
    KSAUDIO_SPEAKER_2POINT1: dict(
        front_left=True,
        front_right=True,
        center=False,
        high_left=False,
        high_right=False,
        side_left=False,
        side_right=False,
        back_left=False,
        back_right=False,
        back_center=False,
        subwoofer=True,
        string='Stereo with subwoofer'
    ),
    KSAUDIO_SPEAKER_3POINT0: dict(
        front_left=True,
        front_right=True,
        center=True,
        high_left=False,
        high_right=False,
        side_left=False,
        side_right=False,
        back_left=False,
        back_right=False,
        back_center=False,
        subwoofer=False,
        string='3.0'
    ),
    KSAUDIO_SPEAKER_3POINT1: dict(
        front_left=True,
        front_right=True,
        center=True,
        high_left=False,
        high_right=False,
        side_left=False,
        side_right=False,
        back_left=False,
        back_right=False,
        back_center=False,
        subwoofer=True,
        string='3.1'
    ),
    KSAUDIO_SPEAKER_QUAD: dict(
        front_left=True,
        front_right=True,
        center=False,
        high_left=False,
        high_right=False,
        side_left=False,
        side_right=False,
        back_left=True,
        back_right=True,
        back_center=False,
        subwoofer=False,
        string='Quad'
    ),
    KSAUDIO_SPEAKER_SURROUND: dict(
        front_left=True,
        front_right=True,
        center=True,
        high_left=False,
        high_right=False,
        side_left=False,
        side_right=False,
        back_left=False,
        back_right=False,
        back_center=True,
        subwoofer=False,
        string='Surround'
    ),
    KSAUDIO_SPEAKER_5POINT0: dict(
        front_left=True,
        front_right=True,
        center=True,
        high_left=False,
        high_right=False,
        side_left=True,
        side_right=True,
        back_left=False,
        back_right=False,
        back_center=False,
        subwoofer=False,
        string='5.0'
    ),
    KSAUDIO_SPEAKER_5POINT1: dict(
        front_left=True,
        front_right=True,
        center=True,
        high_left=False,
        high_right=False,
        side_left=False,
        side_right=False,
        back_left=True,
        back_right=True,
        back_center=False,
        subwoofer=True,
        string='5.1'
    ),
    KSAUDIO_SPEAKER_7POINT0: dict(
        front_left=True,
        front_right=True,
        center=True,
        high_left=False,
        high_right=False,
        side_left=True,
        side_right=True,
        back_left=True,
        back_right=True,
        back_center=False,
        subwoofer=False,
        string='7.0'
    ),
    KSAUDIO_SPEAKER_7POINT1: dict(
        front_left=True,
        front_right=True,
        center=True,
        high_left=True,
        high_right=True,
        side_left=False,
        side_right=False,
        back_left=True,
        back_right=True,
        back_center=False,
        subwoofer=True,
        string='7.1'
    ),
    KSAUDIO_SPEAKER_5POINT1_SURROUND: dict(
        front_left=True,
        front_right=True,
        center=True,
        high_left=False,
        high_right=False,
        side_left=True,
        side_right=True,
        back_left=False,
        back_right=False,
        back_center=False,
        subwoofer=True,
        string='5.1 Surround'
    ),
    KSAUDIO_SPEAKER_7POINT1_SURROUND: dict(
        front_left=True,
        front_right=True,
        center=True,
        high_left=False,
        high_right=False,
        side_left=True,
        side_right=True,
        back_left=True,
        back_right=True,
        back_center=False,
        subwoofer=True,
        string='7.1 Surround'
    )
}

EPCX_PORT_CONNECTION = {
    EPxcPortConnection.ePortConnIntegratedDevice:      'Integrated',
    EPxcPortConnection.ePortConnJack:                  'Physical',
    EPxcPortConnection.ePortConnUnknown:               'Unknown',
    EPxcPortConnection.ePortConnBothIntegratedAndJack: (
        'Integrated and Physical'
    )
}

EPCX_GEN_LOCATION = {
    EPcxGenLocation.eGenLocInternal:   'Internal',
    EPcxGenLocation.eGenLocOther:      'Other',
    EPcxGenLocation.eGenLocPrimaryBox: 'Primary Box',
    EPcxGenLocation.eGenLocSeparate:   'Separate'
}

EPCX_GEO_LOCATION = {
    EPcxGeoLocation.eGeoLocRear:             'Rear',
    EPcxGeoLocation.eGeoLocFront:            'Front',
    EPcxGeoLocation.eGeoLocLeft:             'Left',
    EPcxGeoLocation.eGeoLocRight:            'Right',
    EPcxGeoLocation.eGeoLocTop:              'Top',
    EPcxGeoLocation.eGeoLocBottom:           'Bottom',
    EPcxGeoLocation.eGeoLocRearPanel:        'Rear Panel',
    EPcxGeoLocation.eGeoLocRiser:            'Riser',
    EPcxGeoLocation.eGeoLocInsideMobileLid:  'Inside Mobile Lid',
    EPcxGeoLocation.eGeoLocDrivebay:         'Drive Bay',
    EPcxGeoLocation.eGeoLocHDMI:             'HDMI',
    EPcxGeoLocation.eGeoLocOutsideMobileLid: 'Outside Mobile Lid',
    EPcxGeoLocation.eGeoLocATAPI:            'ATAPI',
    EPcxGeoLocation.eGeoLocNotApplicable:    'Not Applicable',
    EPcxGeoLocation.eGeoLocReserved6:        'Reserved'
}

EPCX_CONNECTION_TYPE = {
    EPcxConnectionType.eConnTypeUnknown:               'Unknown',
    EPcxConnectionType.eConnTypeQuarter:               '6.35mm (1/4" Phono)',
    EPcxConnectionType.eConnTypeAtapiInternal:         'ATAPI',
    EPcxConnectionType.eConnTypeRCA:                   'RCA',
    EPcxConnectionType.eConnTypeOptical:               'Optical',
    EPcxConnectionType.eConnTypeOtherDigital:          'Other Digital',
    EPcxConnectionType.eConnTypeOtherAnalog:           'Analog',
    EPcxConnectionType.eConnTypeXlrProfessional:       'Xlr',
    EPcxConnectionType.eConnTypeRJ11Modem:             'RJ11 (Modem)',
    EPcxConnectionType.eConnTypeCombination:           'Combination',
    EPcxConnectionType.eConnType3Point5mm:             (
        '3.5mm (1\8" Headphone)'
    ),
    EPcxConnectionType.eConnTypeMultichannelAnalogDIN: (
        'Multichannel Analog (DIN)'
    )
}


class JackDescription(object):

    def __init__(self, jack_description):
        self.__jack_description = jack_description

    @property
    def channel_mapping(self):
        return AudioSpeakers(self.__jack_description.ChannelMapping)

    @property
    def color(self):
        return self.__jack_description.Color

    @property
    def type(self):
        return EPCX_CONNECTION_TYPE[self.__jack_description.ConnectionType]

    @property
    def location(self):
        return (
            EPCX_GEN_LOCATION[self.__jack_description.GenLocation] +
            ', ' +
            EPCX_GEO_LOCATION[self.__jack_description.GeoLocation]
        )

    @property
    def port(self):
        return EPCX_PORT_CONNECTION[self.__jack_description.PortConnection]

    @property
    def is_connected(self):
        return bool(self.__jack_description.IsConnected)


class AudioSpeakers(object):

    def __init__(self, value):
        self.value = value

    def __str__(self):
        if self.value is None:
            return 'None'
        return KSAUDIO_SPEAKER[self.value]['string']

    def __getattr__(self, item):
        if item in self.__dict__:
            return self.__dict__[item]

        if self.value is None:
            raise AttributeError

        if item in KSAUDIO_SPEAKER[self.value]:
            return KSAUDIO_SPEAKER[self.value][item]

        raise AttributeError


class _NotificationClient(comtypes.COMObject):
    _com_interfaces_ = [IMMNotificationClient]

    def __init__(self, device_enum, callbacks):
        self.__device_enum = device_enum
        self.__callbacks = callbacks
        comtypes.COMObject.__init__(self)

    def OnDeviceStateChanged(self, pwstrDeviceId, dwNewState):
        device = self.__device_enum.GetDevice(pwstrDeviceId)
        state = AUDIO_DEVICE_STATE[dwNewState]
        for callback in self.__callbacks['state']:
            callback(device, state)

        # print (
        #     'OnDeviceStateChanged',
        #     'pwstrDeviceId:', pwstrDeviceId,
        #     'dwNewState:', dwNewState
        # )

    def OnDeviceAdded(self, pwstrDeviceId):
        device = self.__device_enum.GetDevice(pwstrDeviceId)
        for callback in self.__callbacks['add']:
            callback(device)

        # print 'OnDeviceAdded', 'pwstrDeviceId:', pwstrDeviceId

    def OnDeviceRemoved(self, pwstrDeviceId):
        device = self.__device_enum.GetDevice(pwstrDeviceId)
        for callback in self.__callbacks['remove']:
            callback(device)

        # print 'OnDeviceRemoved', 'pwstrDeviceId:', pwstrDeviceId

    def OnDefaultDeviceChanged(self, _, __, pwstrDefaultDeviceId):
        device = self.__device_enum.GetDevice(pwstrDefaultDeviceId)
        for callback in self.__callbacks['default']:
            callback(device)

        # print (
        #     'OnDefaultDeviceChanged',
        #     'flow:', flow,
        #     'role:', role,
        #     'pwstrDefaultDeviceId:', pwstrDefaultDeviceId
        # )

    def OnPropertyValueChanged(self, pwstrDeviceId, key):
        device = self.__device_enum.GetDevice(pwstrDeviceId)
        for callback in self.__callbacks['property']:
            callback(device, key)
        # print (
        #     'OnPropertyValueChanged',
        #     'pwstrDeviceId:', pwstrDeviceId,
        #     'key:', key
        # )


class AudioDeviceSubunit(object):

    def __init__(self, subunit):
        self.__subunit = subunit

    @property
    def part(self):
        return AudioPart(self.__subunit.QueryInterface(IPart))


class AudioDeviceConnection(object):

    def __init__(self, connector):
        self.__connector = connector

    @property
    def type(self):
        return CONNECTOR_TYPE[self.__connector.GetType().value]

    @property
    def part(self):
        return AudioPart(self.__connector.QueryInterface(IPart))

    @property
    def connected_to_device_id(self):
        return self.__connector.GetDeviceIdConnectedTo()

    @property
    def is_connected(self):
        return self.__connector.IsConnected() == 1

    @property
    def connected_to(self):
        return AudioDeviceConnection(self.__connector.GetConnectedTo())

    @connected_to.setter
    def connected_to(self, connection):
        self.__connector.ConnectTo(connection.get_c_connector())

    @property
    def data_flow(self):
        return E_ROLE[self.__connector.GetDataFlow().value]

    def get_c_connector(self):
        return self.__connector

    def disconnect(self):
        self.__connector.Disconnect()


class AudioPart(object):

    def __init__(self, part):
        self.__part = part

    def activate(self, iid, pointer):
        return comtypes.cast(
            self.__part.Activate(
                CLSCTX_INPROC_SERVER,
                iid
            ),
            pointer
        )

    @property
    def incoming(self):
        return AudioPartsList(self.__part.EnumPartsIncoming())

    @property
    def outgoing(self):
        return AudioPartsList(self.__part.EnumPartsOutgoing())

    @property
    def interfaces(self):
        for index in range(self.__part.GetControlInterfaceCount()):
            yield self.__part.GetControlInterface(index)

    @property
    def global_id(self):
        return self.__part.GetGlobalId()

    @property
    def local_id(self):
        return self.__part.GetLocalId()

    @property
    def name(self):
        return self.__part.GetName()

    @property
    def is_connector(self):
        return self.__part.GetPartType().value == PartType.Connector

    @property
    def is_subunit(self):
        return self.__part.GetPartType().value == PartType.Subunit

    @property
    def sub_type(self):
        return self.__part.GetSubType()

    @property
    def device_topology(self):
        return self.__part.GetTopologyObject()

    def register_control_change_callback(self, interface_iid, callback):
        self.__part.RegisterControlChangeCallback(
            interface_iid,
            callback
        )

    def unregister_control_change_callback(self, callback):
        self.__part.UnregisterControlChangeCallback(callback)
        callback.Release()

    def query_interface(self, interface):
        return self.__part.QueryInterface(interface)


class AudioVolume(object):

    def __init__(self, volume, endpoint):
        self.__volume = volume
        self.__endpoint = endpoint

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

    def up(self):
        self.__volume.VolumeStepUp()

    def down(self):
        self.__volume.VolumeStepDown()

    def peak_meter(self):
        peak_meter = self.__endpoint.activate(
            IID_IAudioMeterInformation,
            PIAudioMeterInformation
        )

        support = peak_meter.QueryHardwareSupport()
        if support | ENDPOINT_HARDWARE_SUPPORT_METER == support:
            return AudioPeakMeter(peak_meter)


class AudioPeakMeter(object):

    def __init__(self, peak_meter):
        self.__peak_meter = peak_meter

    @property
    def channel_peak_values(self):
        return self.__peak_meter.GetChannelsPeakValues(self.channel_count)

    @property
    def channel_count(self):
        return self.__peak_meter.GetMeteringChannelCount()

    @property
    def peak_value(self):
        return self.__peak_meter.GetPeakValue()


class AudioPartsList(object):

    def __init__(self, parts_list):
        self.__parts_list = parts_list

    def get_part(self, index):
        return AudioPart(self.__parts_list.GetPart(index))

    def release(self):
        self.__parts_list.Release()

    def __iter__(self):
        for i in range(self.__parts_list.GetCount()):
            yield self.get_part(i)


class AudioJackSinkInformation(object):

    def __init__(self, sink_information):
        self.__sink_information = (
            sink_information.GetJackSinkInformation()
        )
        sink_information.Release()

    @property
    def manufacturer_id(self):
        return self.__sink_information.ManufacturerId

    @property
    def product_id(self):
        return self.__sink_information.ProductId

    @property
    def audio_latency(self):
        return self.__sink_information.AudioLatency

    @property
    def hdcp_capable(self):
        return bool(self.__sink_information.HDCPCapable)

    @property
    def ai_capable(self):
        return bool(self.__sink_information.AICapable)

    @property
    def description(self):
        return self.__sink_information.SinkDescription

    @property
    def port_id(self):
        return self.__sink_information.PortId

    @property
    def connection_type(self):
        return SINK_CONNECTIONTYPE[self.__sink_information.ConnType.value]


class AudioEndpoint(object):
    def __init__(self, device_enum, endpoint_enum, endpoint):
        self.__device_enum = device_enum
        self.__endpoint_enum = endpoint_enum
        self.__endpoint = endpoint

    @property
    def __device_topology(self):
        return comtypes.cast(
            self.__endpoint.Activate(
                IID_IDeviceTopology,
                CLSCTX_INPROC_SERVER
            ),
            PIDeviceTopology
        )

    @property
    def __connector(self):
        device_topology = self.__device_topology
        return AudioDeviceConnection(device_topology.GetConnector(0))

    @property
    def __subunits(self):
        device_topology = self.__device_topology
        for i in range(device_topology.GetSubunitCount()):
            subunit = device_topology.GetSubunit(i)

            yield AudioDeviceSubunit(
                subunit
            )

    def __activate(self, iid, pointer):
        try:
            return comtypes.cast(
                self.__endpoint.Activate(
                    iid,
                    CLSCTX_INPROC_SERVER
                ),
                pointer
            )
        except comtypes.COMError:
            raise AttributeError

    def __property(self, key):
        pStore = self.__endpoint.OpenPropertyStore(STGM_READ)
        try:
            return pStore.GetValue(key)
        except comtypes.COMError:
            raise AttributeError

    def __get_interface(self, iid, pointer):

        conn_from = self.__connector

        while True:
            try:
                conn_from = conn_from.connected_to
            except comtypes.COMError:
                return

            part = conn_from.part
            device_topology = part.device_topology

            for i in range(device_topology.GetSubunitCount()):
                subunit = AudioDeviceSubunit(device_topology.GetSubunit(i))
                part = subunit.part
                try:
                    interface = part.activate(iid, pointer)
                    print interface
                    return interface
                except comtypes.COMError:
                    continue

            if conn_from.type == 'Software IO':
                raise AttributeError

            if not conn_from.is_connected:
                raise AttributeError

    @property
    def audio_client(self):
        return self.__activate(
            IID_IAudioClient,
            PIAudioClient
        )

    @property
    def session_manager(self):
        return self.__activate(
            IID_IAudioSessionManager2,
            PIAudioSessionManager2
        )

    @property
    def data_flow(self):
        try:
            endpoint = self.__endpoint.QueryInterface(IMMEndpoint)
            return E_DATA_FLOW[endpoint.GetDataFlow().value]
        except comtypes.COMError:
            raise AttributeError

    @property
    def name(self):
        """Return an endpoint devices FriendlyName."""
        return self.__property(PKEY_Device_FriendlyName)

    @property
    def description(self):
        return self.__property(PKEY_Device_DeviceDesc)

    @property
    def form_factor(self):
        form_factor = self.__property(PKEY_AudioEndpoint_FormFactor)
        return ENDPOINT_FORM_FACTOR[form_factor]

    @property
    def type(self):
        return self.__property(PKEY_AudioEndpoint_JackSubType)

    @property
    def full_range_speakers(self):
        return AudioSpeakers(
            self.__property(PKEY_AudioEndpoint_FullRangeSpeakers)
        )

    @property
    def guid(self):
        return self.__property(PKEY_AudioEndpoint_GUID)

    @property
    def physical_speakers(self):
        return AudioSpeakers(
            self.__property(PKEY_AudioEndpoint_PhysicalSpeakers)
        )

    @property
    def system_effects(self):
        return bool(
            self.__property(PKEY_AudioEndpoint_Disable_SysFx)
        )

    @property
    def jack_descriptions(self):
        conn_from = self.__connector
        try:
            conn_to = conn_from.connected_to
        except comtypes.COMError:
            raise AttributeError
        part = conn_to.part

        try:
            jack_description = part.activate(
                IID_IKsJackDescription,
                PIKsJackDescription
            )
        except comtypes.COMError:
            raise AttributeError

        for i in range(jack_description.GetJackCount()):
            yield JackDescription(jack_description.GetJackDescription(i))

    @property
    def jack_information(self):
        raise NotImplementedError
        # conn_from = self.__connector
        # try:
        #     conn_to = conn_from.connected_to
        # except comtypes.COMError:
        #     raise AttributeError
        #
        # part = conn_to.part
        #
        # try:
        #     return AudioJackSinkInformation(
        #         part.activate(
        #             IID_IKsJackSinkInformation,
        #             PIKsJackSinkInformation
        #         )
        #     )
        # except comtypes.COMError:
        #     raise AttributeError

    @property
    def auto_gain_control(self):
        return self.__get_interface(
            IID_IAudioAutoGainControl,
            PIAudioAutoGainControl
        )

    @property
    def bass(self):
        return self.__get_interface(
            IID_IAudioBass,
            PIAudioBass
        )

    @property
    def channel_config(self):
        return self.__get_interface(
            IID_IAudioChannelConfig,
            PIAudioChannelConfig
        )

    @property
    def input(self):
        return self.__get_interface(
            IID_IAudioInputSelector,
            PIAudioInputSelector
        )

    @property
    def loudness(self):
        return self.__get_interface(
            IID_IAudioLoudness,
            PIAudioLoudness
        )

    @property
    def midrange(self):
        return self.__get_interface(
            IID_IAudioMidrange,
            PIAudioMidrange
        )

    @property
    def output(self):
        return self.__get_interface(
            IID_IAudioOutputSelector,
            PIAudioOutputSelector
        )

    @property
    def peak_meter(self):
        return self.__get_interface(
            IID_IAudioPeakMeter,
            PIAudioPeakMeter
        )

    @property
    def treble(self):
        return self.__get_interface(
            IID_IAudioTreble,
            PIAudioTreble
        )

    @property
    def volume(self):
        volume = self.__activate(
            IID_IAudioEndpointVolumeEx,
            PIAudioEndpointVolumeEx
        )
        support = volume.QueryHardwareSupport()
        if support | ENDPOINT_HARDWARE_SUPPORT_VOLUME == support:
            return AudioVolume(volume, self.__endpoint)
        raise AttributeError

    @property
    def format_support(self):
        raise NotImplementedError
        # if str(IID_IKsFormatSupport) == str(interface.GetIID()):
        #     part = self.__connector.QueryInterface(IPart)
        #
        #     control = comtypes.cast(
        #         part.Activate(
        #             comtypes.CLSCTX_INPROC_SERVER,
        #             IID_IKsFormatSupport
        #         ),
        #         PIKsFormatSupport
        #     )
        #     print 'IKsFormatSupport:', control
        #
        #     return None


class AudioDevice(object):

    def __init__(self, dev_id, device_enum, data_flow):

        self.__device_enum = device_enum
        self.__id = dev_id
        self.__data_flow = data_flow

    @property
    def __device(self):
        return self.__device_enum.GetDevice(self.__id)

    @property
    def __device_topology(self):
        return comtypes.cast(
            self.__device.Activate(
                IID_IDeviceTopology,
                CLSCTX_INPROC_SERVER
            ),
            PIDeviceTopology
        )

    @property
    def __connectors(self):
        name = self.name
        endpoint_enum = self.__device_enum.EnumAudioEndpoints(
            self.__data_flow,
            DEVICE_STATE_MASK_ALL
        )

        for i in range(endpoint_enum.GetCount()):
            endpoint = endpoint_enum.Item(i)

            pStore = endpoint.OpenPropertyStore(STGM_READ)
            try:
                item_name = pStore.GetValue(PKEY_DeviceInterface_FriendlyName)
            except comtypes.COMError:
                continue

            if item_name == name:
                for j in range(self.__device_topology.GetConnectorCount()):
                    yield AudioDeviceConnection(
                        self.__device_topology.GetConnector(j)
                    )

    @property
    def __subunits(self):
        for i in range(self.__device_topology.GetSubunitCount()):
            subunit = self.__device_topology.GetSubunit(i)
            yield subunit.QueryInterface(IPart)

    @property
    def connectors(self):
        return self.__connectors

    @property
    def state(self):
        return AUDIO_DEVICE_STATE[self.__device.GetState()]

    @property
    def connector_count(self):
        return self.__device_topology.GetConnectorCount()

    @property
    def id(self):
        return self.__device_topology.GetDeviceId()

    @property
    def name(self):
        pStore = self.__device.OpenPropertyStore(STGM_READ)
        try:
            return pStore.GetValue(PKEY_DeviceInterface_FriendlyName)
        except comtypes.COMError:
            pass

    @property
    def render_endpoints(self):
        return list(self.__endpoints(EDataFlow.eRender))

    @property
    def capture_endpoints(self):
        return list(self.__endpoints(EDataFlow.eCapture))

    def __endpoints(self, data_flow):
        endpoint_enum = self.__device_enum.EnumAudioEndpoints(
            data_flow,
            DEVICE_STATE_MASK_ALL
        )
        device_name = self.name

        for i in range(endpoint_enum.GetCount()):
            endpoint = endpoint_enum.Item(i)

            pStore = endpoint.OpenPropertyStore(STGM_READ)
            try:
                name = pStore.GetValue(PKEY_DeviceInterface_FriendlyName)
            except comtypes.COMError:
                continue

            if name == device_name:
                yield AudioEndpoint(
                    self.__device_enum,
                    endpoint_enum,
                    endpoint
                )

    def __iter__(self):
        return iter(self.__endpoints(EDataFlow.eAll))


class AudioDevices(object):

    def __init__(self):
        comtypes.CoInitialize()

        self.__callbacks = dict(
            property=[],
            state=[],
            default=[],
            remove=[],
            add=[]
        )

        self.__device_enum = comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator,
            IMMDeviceEnumerator,
            comtypes.CLSCTX_INPROC_SERVER
        )

        self.__notification_client = _NotificationClient(
            self.__device_enum,
            self.__callbacks
        )
        self.__device_enum.RegisterEndpointNotificationCallback(
            self.__notification_client
        )

    def bind(self, notification, callback):
        self.__callbacks[notification] += callback

    def unbind(self, notification, callback):
        self.__callbacks[notification].remove(callback)

    @property
    def render_devices(self):
        return list(self.__devices(EDataFlow.eRender))

    @property
    def capture_devices(self):
        return list(self.__devices(EDataFlow.eCapture))

    def __devices(self, data_flow):
        endpoint_enum = self.__device_enum.EnumAudioEndpoints(
            data_flow,
            DEVICE_STATE_MASK_ALL
        )

        used_names = []

        for i in range(endpoint_enum.GetCount()):
            endpoint = endpoint_enum.Item(i)

            pStore = endpoint.OpenPropertyStore(STGM_READ)
            try:
                name = pStore.GetValue(PKEY_DeviceInterface_FriendlyName)
            except comtypes.COMError:
                continue

            if name not in used_names:
                used_names += [name]
                device_topology = comtypes.cast(
                    endpoint.Activate(
                        IID_IDeviceTopology,
                        CLSCTX_INPROC_SERVER
                    ),
                    PIDeviceTopology
                )

                yield AudioDevice(
                    device_topology.GetDeviceId(),
                    self.__device_enum,
                    data_flow
                )

    def __iter__(self):
        for device in self.__devices(EDataFlow.eAll):
            yield device

    def __getitem__(self, item):
        if isinstance(item, (slice, int)):
            return list(self)[item]

        for device in self:
            if item in (device.name, device.id):
                return device
        raise AttributeError
