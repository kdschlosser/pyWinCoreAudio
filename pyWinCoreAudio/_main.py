
import threading
import comtypes
import ctypes
from _audioclient import PIAudioClient
from _audiopolicy import (
    IAudioSessionEvents,
    IAudioSessionNotification,
    IAudioSessionControl2,
    PIAudioSessionManager,
    PIAudioSessionManager2
)
from _policyconfig import (
    IPolicyConfigVista,
)

from _endpointvolumeapi import (
    PIAudioEndpointVolumeEx,
    PIAudioMeterInformation,
    IAudioEndpointVolumeCallback
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
    EChannelMapping,
    AudioSessionState,
    AudioSessionDisconnectReason,
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
    IID_IAudioSessionManager,
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
    CLSID_PolicyConfigVistaClient
)
from _constant import (
    S_OK,
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
    # ENDPOINT_HARDWARE_SUPPORT_METER,
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
    JACKDESC2_PRESENCE_DETECT_CAPABILITY,
    JACKDESC2_DYNAMIC_FORMAT_CHANGE_CAPABILITY
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


def convert_triplet_to_rgb(triplet):
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

CHANNEL_MAPPING = {
    EChannelMapping.ePcxChanMap_FL_FR: (
        KSAUDIO_SPEAKER[KSAUDIO_SPEAKER_STEREO]
    ),
    EChannelMapping.ePcxChanMap_FC_LFE: (
        KSAUDIO_SPEAKER[KSAUDIO_SPEAKER_1POINT1]
    ),
    EChannelMapping.ePcxChanMap_FLC_FRC: (
        KSAUDIO_SPEAKER[KSAUDIO_SPEAKER_3POINT0]
    ),
    EChannelMapping.ePcxChanMap_BL_BR: dict(
        front_left=False,
        front_right=False,
        center=False,
        high_left=False,
        high_right=False,
        side_left=False,
        side_right=False,
        back_left=True,
        back_right=True,
        back_center=False,
        subwoofer=False,
        string='Back left & right'
    ),

    EChannelMapping.ePcxChanMap_SL_SR: dict(
        front_left=False,
        front_right=False,
        center=False,
        high_left=False,
        high_right=False,
        side_left=True,
        side_right=True,
        back_left=False,
        back_right=False,
        back_center=False,
        subwoofer=False,
        string='Side left & right'
    ),
    EChannelMapping.ePcxChanMap_Unknown: dict(
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
        string='Unknown'
    ),
}

EPCX_PORT_CONNECTION = {
    EPxcPortConnection.ePortConnJack:                  'Jack',
    EPxcPortConnection.ePortConnUnknown:               'Unknown',
    EPxcPortConnection.ePortConnIntegratedDevice:      (
        'Slot for an integrated device'
    ),
    EPxcPortConnection.ePortConnBothIntegratedAndJack: (
        'Both a jack and a slot for an integrated device'
    )
}

EPCX_GEN_LOCATION = {
    EPcxGenLocation.eGenLocInternal:   'Inside primary chassis',
    EPcxGenLocation.eGenLocOther:      'Other location',
    EPcxGenLocation.eGenLocPrimaryBox: 'On primary chassis',
    EPcxGenLocation.eGenLocSeparate:   'On separate chassis'
}

EPCX_GEO_LOCATION = {
    EPcxGeoLocation.eGeoLocRear:             'Rear-mounted panel',
    EPcxGeoLocation.eGeoLocFront:            'Front-mounted panel',
    EPcxGeoLocation.eGeoLocLeft:             'Left-mounted panel',
    EPcxGeoLocation.eGeoLocRight:            'Right-mounted panel',
    EPcxGeoLocation.eGeoLocTop:              'Top-mounted panel',
    EPcxGeoLocation.eGeoLocBottom:           'Bottom-mounted panel',
    EPcxGeoLocation.eGeoLocRiser:            'Riser card',
    EPcxGeoLocation.eGeoLocInsideMobileLid:  'Inside lid of mobile computer',
    EPcxGeoLocation.eGeoLocDrivebay:         'Drive bay',
    EPcxGeoLocation.eGeoLocHDMI:             'HDMI connector',
    EPcxGeoLocation.eGeoLocOutsideMobileLid: 'Outside lid of mobile computer',
    EPcxGeoLocation.eGeoLocATAPI:            'ATAPI connector',
    EPcxGeoLocation.eGeoLocNotApplicable:    'Not Applicable',
    EPcxGeoLocation.eGeoLocReserved6:        'Reserved',
    EPcxGeoLocation.eGeoLocRearPanel:        (
        'Rear slide-open or pull-open panel'
    ),
    EPcxGeoLocation.eGeoLocRearOPanel:       (
        'Rear slide-open or pull-open panel'
    )
}

EPCX_CONNECTION_TYPE = {
    EPcxConnectionType.eConnTypeUnknown:               'Unknown',
    EPcxConnectionType.eConnTypeRCA:                   'RCA jack',
    EPcxConnectionType.eConnTypeOptical:               'Optical connector',
    EPcxConnectionType.eConnTypeXlrProfessional:       'XLR connector',
    EPcxConnectionType.eConnTypeRJ11Modem:             'RJ11 modem connector',
    EPcxConnectionType.eConnTypeQuarter:               (
        '6.35mm (1/4" Phono) jack'
    ),
    EPcxConnectionType.eConnTypeOtherDigital:          (
        'Generic digital connector'
    ),
    EPcxConnectionType.eConnTypeOtherAnalog:           (
        'Generic analog connector'
    ),
    EPcxConnectionType.eConnTypeAtapiInternal:         (
        'ATAPI internal connector'
    ),
    EPcxConnectionType.eConnTypeCombination:           (
        'Combination of connector types'
    ),
    EPcxConnectionType.eConnType3Point5mm:             (
        '3.5mm (1\8" Headphone) jack'
    ),
    EPcxConnectionType.eConnTypeEighth:                (
        '3.5mm (1\8" Headphone) jack'
    ),
    EPcxConnectionType.eConnTypeMultichannelAnalogDIN: (
        'Multichannel analog DIN connector'
    )
}

AUDIO_SESSION_STATE = {
    AudioSessionState.AudioSessionStateInactive: 'Inactive',
    AudioSessionState.AudioSessionStateActive:   'Active',
    AudioSessionState.AudioSessionStateExpired:  'Expired'
}


AUDIO_SESSION_DISCONNECT_REASON = {
    AudioSessionDisconnectReason.DisconnectReasonDeviceRemoval:         (
        'Audio endpoint device removed'
    ),
    AudioSessionDisconnectReason.DisconnectReasonServerShutdown:        (
        'Windows audio service has stopped'
    ),
    AudioSessionDisconnectReason.DisconnectReasonFormatChanged:         (
        'Stream format changed'
    ),
    AudioSessionDisconnectReason.DisconnectReasonSessionLogoff:         (
        'User logged off'
    ),
    AudioSessionDisconnectReason.DisconnectReasonSessionDisconnected:   (
        'Remote desktop session disconnected'
    ),
    AudioSessionDisconnectReason.DisconnectReasonExclusiveModeOverride: (
        'Shared mode disabled'
    ),

}


class Singleton(type):
    _instances = {}

    def __call__(cls, *args):
        if cls not in cls._instances:
            cls._instances[cls] = {}

        instances = cls._instances[cls]

        for key, value in instances.items():
            if key == args:
                return value

        instances[args] = super(
            Singleton,
            cls
        ).__call__(*args)

        # print '{'
        #
        # for key, value in cls._instances.items():
        #     print '   ', repr(str(key)) + ': {'
        #     for k, v in value.items():
        #         k = ', '.join(repr(str(item)) for item in k)
        #         print '        (' + k + '):', repr(str(v)) + ','
        #     print '    },'
        # print '}'

        return instances[args]


def _run_in_thread(func, *args, **kwargs):
    t = threading.Thread(target=func, args=args, kwargs=kwargs)
    t.daemon = True
    t.start()
    return t


class JackDescription(object):

    def __init__(self, jack_description, jack_description2):
        self.__jack_description = jack_description
        self.__jack_description2 = jack_description2

    @property
    def channel_mapping(self):
        return AudioSpeakers(self.__jack_description.ChannelMapping)

    @property
    def color(self):
        return convert_triplet_to_rgb(self.__jack_description.Color)

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

    @property
    def presence_detection(self):
        if self.__jack_description2 is None:
            raise AttributeError

        return (
            self.__jack_description2.JackCapabilities |
            JACKDESC2_PRESENCE_DETECT_CAPABILITY ==
            self.__jack_description2.JackCapabilities
        )

    @property
    def dynamic_format_change(self):
        if self.__jack_description2 is None:
            raise AttributeError

        return (
            self.__jack_description2.JackCapabilities |
            JACKDESC2_DYNAMIC_FORMAT_CHANGE_CAPABILITY ==
            self.__jack_description2.JackCapabilities
        )


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


class AudioEndpointVolumeCallback(comtypes.COMObject):
    _com_interfaces_ = [IAudioEndpointVolumeCallback]

    def __init__(self, endpoint, callback):
        self.__endpoint = endpoint
        self.__callback = callback
        comtypes.COMObject.__init__(self)

    def OnNotify(self, pNotify):
        mute = bool(pNotify.contents.bMuted)
        master_volume = pNotify.contents.fMasterVolume
        num_channels = pNotify.contents.nChannels
        pfChannelVolumes = ctypes.cast(
            pNotify.contents.afChannelVolumes,
            ctypes.POINTER(ctypes.c_float)
        )
        channel_volumes = list(
            pfChannelVolumes[i] for i in range(num_channels)
        )

        def do():
            self.__callback.endpoint_volume_change(
                self.__endpoint,
                master_volume,
                channel_volumes,
                mute
            )

        _run_in_thread(do)

        return S_OK


class AudioNotificationClient(comtypes.COMObject):
    _com_interfaces_ = [IMMNotificationClient]

    def __init__(self, device_enum, callbacks):
        self.__device_enum = device_enum
        self.__callbacks = callbacks
        comtypes.COMObject.__init__(self)

    def __get_device(self, device_id):
        device = self.__device_enum.get_device(device_id)
        device_topology = comtypes.cast(
            device.Activate(
                IID_IDeviceTopology,
                CLSCTX_INPROC_SERVER
            ),
            PIDeviceTopology
        )

        return AudioDevice(
            device_topology.GetDeviceId(),
            self.__device_enum
        )

    def OnDeviceStateChanged(self, pwstrDeviceId, dwNewState):
        device = self.__get_device(pwstrDeviceId)
        state = AUDIO_DEVICE_STATE[dwNewState]

        def do():
            for callback in self.__callbacks:
                try:
                    callback.device_state_change(device, state)
                except AttributeError:
                    pass

        _run_in_thread(do)

        return S_OK

    def OnDeviceAdded(self, pwstrDeviceId):
        device = self.__get_device(pwstrDeviceId)

        def do():
            for callback in self.__callbacks:
                try:
                    callback.device_added(device)
                except AttributeError:
                    pass

        _run_in_thread(do)

        return S_OK

    def OnDeviceRemoved(self, pwstrDeviceId):
        device = self.__get_device(pwstrDeviceId)

        def do():
            for callback in self.__callbacks:
                try:
                    callback.device_removed(device)
                except AttributeError:
                    pass

        _run_in_thread(do)

        return S_OK

    def OnDefaultDeviceChanged(self, _, __, pwstrDefaultDeviceId):
        device = self.__get_device(pwstrDefaultDeviceId)

        def do():
            for callback in self.__callbacks:
                try:
                    callback.default_endpoint_changed(device)
                except AttributeError:
                    pass

        _run_in_thread(do)

        return S_OK

    def OnPropertyValueChanged(self, pwstrDeviceId, key):
        device = self.__get_device(pwstrDeviceId)

        def do():
            for callback in self.__callbacks:
                try:
                    callback.device_property_changed(device, key)
                except AttributeError:
                    pass

        _run_in_thread(do)

        return S_OK


class AudioWaveFormat(object):

    def __init__(self, client):
        self.__client = client


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

    def register_notification_callback(self, interface_iid, callback):
        self.__part.RegisterControlChangeCallback(
            interface_iid,
            callback
        )

    def unregister_notification_callback(self, callback):
        self.__part.UnregisterControlChangeCallback(callback)
        callback.Release()

    def query_interface(self, interface):
        return self.__part.QueryInterface(interface)


class AudioVolume(object):
    __metaclass__ = Singleton

    def __init__(self, endpoint):
        self.__endpoint = endpoint
        self.__volume = endpoint.activate(
            IID_IAudioEndpointVolumeEx,
            PIAudioEndpointVolumeEx
        )
        support = self.__volume.QueryHardwareSupport()
        if support | ENDPOINT_HARDWARE_SUPPORT_VOLUME != support:
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
            PIAudioMeterInformation
        )

    @property
    def channel_peak_values(self):

        # support = self.__peak_meter.QueryHardwareSupport()
        # if support | ENDPOINT_HARDWARE_SUPPORT_METER != support:
            # raise NotImplementedError
        channels = self.__peak_meter.GetChannelsPeakValues(self.channel_count)

        channel_peaks = ctypes.cast(
            channels,
            ctypes.POINTER(ctypes.c_float)
        )
        return list(
            channel_peaks[i] for i in range(self.channel_count)
        )

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


class AudioSessionNotification(comtypes.COMObject):
    _com_interfaces_ = [IAudioSessionNotification]

    def __init__(self, session_manager, callback):
        self.__session_manager = session_manager
        self.__callback = callback
        comtypes.COMObject.__init__(self)

    def OnSessionCreated(self, _):

        def do():
            self.__session_manager.update_sessions(
                self.__callback.session_created
            )

        _run_in_thread(do)

        return S_OK


class AudioSessionManager(object):
    def __init__(self, endpoint):
        self.__endpoint = endpoint
        self.__sessions = {}
        self.__update_lock = threading.Lock()

        try:
            self.__session_manager = endpoint.activate(
                IID_IAudioSessionManager2,
                PIAudioSessionManager2
            )
        except comtypes.COMError:
            try:
                self.__session_manager = endpoint.activate(
                    IID_IAudioSessionManager,
                    PIAudioSessionManager
                )
            except comtypes.COMError:
                raise NotImplementedError

    def update_sessions(self, callback=None):

        self.__update_lock.acquire()

        event = threading.Event()
        event.wait(0.1)

        try:
            try:
                session_enum = self.__session_manager.GetSessionEnumerator()
            except comtypes.COMError:
                pass

            else:
                sessions = {}

                for i in range(session_enum.GetCount()):
                    session = session_enum.GetSession(i)
                    try:
                        session = session.QueryInterface(IAudioSessionControl2)
                    except comtypes.COMError:
                        pass

                    if session in self.__sessions:
                        sessions[session] = self.__sessions[session]
                        del self.__sessions[session]

                    else:
                        sessions[session] = AudioSession(self, session)
                        if callback is not None:
                            callback(sessions[session])

                for session in self.__sessions.values():
                    session.release()

                self.__sessions.clear()

                for key, value in sessions.items():
                    self.__sessions[key] = value
        finally:
            self.__update_lock.release()

    @property
    def endpoint(self):
        return self.__endpoint

    def register_duck_notification(self, callback):
        raise NotImplementedError

    def unregister_duck_notification(self, callback):
        raise NotImplementedError

    def register_notification_callback(self, callback):
        try:
            callback = AudioSessionNotification(self, callback)
            self.__session_manager.RegisterSessionNotification(callback)
            return callback
        except comtypes.COMError:
            raise AttributeError

    def unregister_notification_callback(self, callback):
        try:
            self.__session_manager.UnregisterSessionNotification(callback)
        except comtypes.COMError:
            raise AttributeError

    def __iter__(self):
        try:
            session_enum = self.__session_manager.GetSessionEnumerator()
        except comtypes.COMError:
            raise TypeError

        for i in range(session_enum.GetCount()):
            session = session_enum.GetSession(i)
            try:
                session = session.QueryInterface(IAudioSessionControl2)
            except comtypes.COMError:
                pass

            if session not in self.__sessions:
                self.__sessions[session] = AudioSession(self, session)

            yield self.__sessions[session]


class AudioSessionEvent(comtypes.COMObject):
    _com_interfaces_ = [IAudioSessionEvents]

    def __init__(self, session, callback):
        self.__session = session
        self.__callback = callback
        self.__name = session.name
        self.__endpoint = session.session_manager.endpoint

        try:
            self.__id = session.id
        except AttributeError:
            self.__id = None

        comtypes.COMObject.__init__(self)

    def OnChannelVolumeChanged(
        self,
        ChannelCount,
        NewChannelVolumeArray,
        ChangedChannel,
        _
    ):
        channel_volume_array = ctypes.cast(
            NewChannelVolumeArray,
            ctypes.POINTER(ctypes.c_float)
        )
        channel_volumes = list(
            channel_volume_array[i] for i in range(ChannelCount)
        )

        def do():
            self.__callback.channel_volume_changed(
                self.__session,
                ChangedChannel,
                channel_volumes[ChangedChannel]
            )

        _run_in_thread(do)

        return S_OK

    def OnDisplayNameChanged(self, NewDisplayName, _):
        if NewDisplayName == '@%SystemRoot%\System32\AudioSrv.Dll,-202':
            NewDisplayName = 'System Sounds'

        def do():
            self.__callback.session_name_changed(
                self.__session,
                NewDisplayName,
            )

        _run_in_thread(do)

        return S_OK

    def OnGroupingParamChanged(self, NewGroupingParam, _):

        def do():
            self.__callback.session_grouping_changed(
                self.__session,
                NewGroupingParam,
            )

        _run_in_thread(do)

        return S_OK

    def OnIconPathChanged(self, NewIconPath, _):
        def do():
            self.__callback.session_icon_path_changed(
                self.__session,
                NewIconPath,
            )

        _run_in_thread(do)

        return S_OK

    def OnSessionDisconnected(self, DisconnectReason):

        def do():
            self.__callback.session_disconnect(
                self.__endpoint,
                self.__name,
                self.__id,
                AUDIO_SESSION_DISCONNECT_REASON[DisconnectReason]
            )
            self.__session.session_manager.update_sessions()

        _run_in_thread(do)

        return S_OK

    def OnSimpleVolumeChanged(self, NewVolume, NewMute, _):

        def do():
            self.__callback.session_volume_changed(
                self.__session,
                NewVolume,
                bool(NewMute)
            )

        _run_in_thread(do)

        return S_OK

    def OnStateChanged(self, NewState):

        def do():
            self.__callback.session_state_changed(
                self.__session,
                AUDIO_SESSION_STATE[NewState]
            )

        _run_in_thread(do)

        return S_OK


class AudioSession(object):

    def __init__(self, session_manager, session):
        self.__session_manager = session_manager
        self.__session = session

    def release(self):
        del self.__session

    @property
    def session_manager(self):
        return self.__session_manager

    @property
    def name(self):
        display_name = self.__session.GetDisplayName()
        if display_name == '@%SystemRoot%\System32\AudioSrv.Dll,-202':
            display_name = 'System Sounds'
        return display_name

    @name.setter
    def name(self, name):
        self.__session.SetDisplayName(name, None)

    @property
    def grouping_param(self):
        return self.__session.GetGroupingParam()

    @grouping_param.setter
    def grouping_param(self, grouping_param):
        self.__session.SetGroupingParam(grouping_param, None)

    @property
    def icon_path(self):
        return self.__session.GetIconPath()

    @icon_path.setter
    def icon_path(self, icon_path):
        self.__session.SetIconPath(icon_path, None)

    @property
    def state(self):
        return AUDIO_SESSION_STATE[self.__session.GetState().value]

    @property
    def process_id(self):
        try:
            return self.__session.GetProcessId()
        except comtypes.COMError:
            raise AttributeError

    @property
    def id(self):
        try:
            return self.__session.GetSessionIdentifier()
        except comtypes.COMError:
            raise AttributeError

    @property
    def instance_id(self):
        try:
            return self.__session.GetSessionInstanceIdentifier()
        except comtypes.COMError:
            raise AttributeError

    @property
    def is_system_sounds(self):
        try:
            return not bool(self.__session.IsSystemSoundsSession())
        except comtypes.COMError:
            raise AttributeError

    def ducking_preferences(self, ducking_preference):
        try:
            self.__session.SetDuckingPreference(ducking_preference)
        except comtypes.COMError:
            raise AttributeError

    ducking_preferences = property(fset=ducking_preferences)

    def register_notification_callback(self, callback):
        callback = AudioSessionEvent(self, callback)
        self.__session.RegisterAudioSessionNotification(callback)
        return callback

    def unregister_notification_callback(self, callback):
        self.__session.UnregisterAudioSessionNotification(callback)


class AudioEndpoint(object):
    __metaclass__ = Singleton

    def __init__(
        self,
        parent,
        endpoint_id,
        device_enum
    ):
        self.__parent = parent
        self.__endpoint_id = endpoint_id
        self.__device_enum = device_enum
        self.__endpoint = self.__device_enum.get_device(endpoint_id)
        self.activate = self.__activate

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

    def set_default(self):

        self.__device_enum = comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator,
            IMMDeviceEnumerator,
            comtypes.CLSCTX_INPROC_SERVER
        )

        policy_config = comtypes.CoCreateInstance(
            CLSID_PolicyConfigVistaClient,
            IPolicyConfigVista,
            comtypes.CLSCTX_ALL
        )

        policy_config.SetDefaultEndpoint(self.id, ERole.eMultimedia)

    @property
    def device(self):
        return self.__parent

    @property
    def id(self):
        return self.__endpoint_id

    @property
    def audio_client(self):
        return self.__activate(
            IID_IAudioClient,
            PIAudioClient
        )

    @property
    def session_manager(self):
        try:
            return AudioSessionManager(self)
        except NotImplementedError:
            raise AttributeError

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

        try:
            jack_description2 = part.activate(
                IID_IKsJackDescription2,
                PIKsJackDescription2
            )
        except comtypes.COMError:
            jack_description2 = None

        for i in range(jack_description.GetJackCount()):
            jd = jack_description.GetJackDescription(i)
            if jack_description2 is None:
                jd2 = None
            else:
                jd2 = jack_description2.GetJackDescription2(i)
            yield JackDescription(jd, jd2)

    @property
    def jack_information(self):
        conn_from = self.__connector
        try:
            conn_to = conn_from.connected_to
        except comtypes.COMError:
            raise AttributeError

        part = conn_to.part

        try:
            return AudioJackSinkInformation(
                part.activate(
                    IID_IKsJackSinkInformation,
                    PIKsJackSinkInformation
                )
            )
        except comtypes.COMError:
            raise AttributeError

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
        try:
            return AudioVolume(self)
        except NotImplementedError:
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

    @property
    def is_default(self):
        return AudioDefaultEndpoint(self.__device_enum, self.data_flow) == self


class AudioDevice(object):
    __metaclass__ = Singleton

    def __init__(self, dev_id, device_enum):
        self.__id = dev_id
        self.__device_enum = device_enum

    @property
    def __device(self):
        return self.__device_enum.get_device(self.__id)

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
        for endpoint in self.__device_enum.endpoints:

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
        endpoint_enum = self.__device_enum.endpoint_enum(data_flow)
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
                    self,
                    endpoint.GetId(),
                    self.__device_enum
                )

    def __iter__(self):
        return iter(self.__endpoints(EDataFlow.eAll))


class AudioDefaultEndpoint(object):

    def __init__(self, device_enum, data_flow):
        self.__device_enum = device_enum
        for key, value in E_DATA_FLOW.items():
            if data_flow in (key, value):
                data_flow = key
                break
        else:
            raise AttributeError

        self.__data_flow = data_flow

    @property
    def data_flow(self):
        return self.__data_flow

    @property
    def __default_endpoint(self):
        endpoint = self.__device_enum.default_endpoint(self.__data_flow)

        device_topology = comtypes.cast(
            endpoint.Activate(
                IID_IDeviceTopology,
                CLSCTX_INPROC_SERVER
            ),
            PIDeviceTopology
        )

        device = AudioDevice(
            device_topology.GetDeviceId(),
            self.__device_enum
        )

        endpoint_id = endpoint.GetId()

        for endpoint in device:
            if endpoint.id == endpoint_id:
                return endpoint

    def __eq__(self, other):
        return other == self.__default_endpoint

    def __getattr__(self, item):
        if item in self.__dict__:
            return self.__dict__[item]

        endpoint = self.__default_endpoint
        if endpoint is not None:
            return getattr(endpoint, item)

        raise AttributeError


class AudioDeviceEnumerator(object):

    def __init__(self):
        self.__device_enum = comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator,
            IMMDeviceEnumerator,
            comtypes.CLSCTX_INPROC_SERVER
        )

    def register_endpoint_notification_callback(self, client):
        self.__device_enum.RegisterEndpointNotificationCallback(
            client
        )

    def unregister_endpoint_notification_callback(self, client):
        self.__device_enum.UnregisterEndpointNotificationCallback(
            client
        )

    def default_endpoint(self, data_flow):

        return self.__device_enum.GetDefaultAudioEndpoint(
            data_flow,
            ERole.eMultimedia
        )

    def get_device(self, device_id):
        return self.__device_enum.GetDevice(device_id)

    @property
    def endpoints(self):
        endpoint_enum = self.__device_enum.EnumAudioEndpoints(
            EDataFlow.eAll,
            DEVICE_STATE_MASK_ALL
        )

        for i in range(endpoint_enum.GetCount()):
            yield endpoint_enum.Item(i)

    def endpoint_enum(self, data_flow):
        return self.__device_enum.EnumAudioEndpoints(
            data_flow,
            DEVICE_STATE_MASK_ALL
        )


class AudioDevices(object):

    def __init__(self):
        comtypes.CoInitialize()

        self.__callbacks = []
        self.__device_enum = AudioDeviceEnumerator()

        self.__notification_client = AudioNotificationClient(
            self.__device_enum,
            self.__callbacks
        )
        self.__device_enum.register_endpoint_notification_callback(
            self.__notification_client
        )

    def register_notification_callback(self, callback):
        self.__callbacks += [callback]

    def unregister_notification_callback(self, callback):
        self.__callbacks.remove(callback)

    @property
    def default_render_endpoint(self):
        return AudioDefaultEndpoint(self.__device_enum, EDataFlow.eRender)

    @property
    def default_capture_endpoint(self):
        return AudioDefaultEndpoint(self.__device_enum, EDataFlow.eCapture)

    @property
    def __devices(self):
        used_names = []

        for endpoint in self.__device_enum.endpoints:
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
                    self.__device_enum
                )

    def __iter__(self):
        for device in self.__devices:
            yield device

    def __contains__(self, item):
        for device in self:
            if item in (device.name, device.id):
                return True
        return False

    def __getitem__(self, item):
        if isinstance(item, (slice, int)):
            return list(self)[item]

        for device in self:
            if item in (device.name, device.id):
                return device
        raise AttributeError
