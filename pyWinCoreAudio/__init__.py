
import comtypes
from _endpointvolumeapi import (
    IAudioEndpointVolumeEx,
    PIAudioEndpointVolume,
    PIAudioMeterInformation,
)
from _audioclient import PIAudioClient
from _iid import (
    IID_IDeviceTopology,
    IID_IMMEndpoint,
    IID_IAudioEndpointVolume,
    IID_IPart,
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
    IID_IDeviceSpecificProperty,
    CLSID_MMDeviceEnumerator,
    CLSID_DeviceTopology,
    IID_IKsFormatSupport,
    IID_IKsJackDescription2,
    IID_IKsControl,
    IID_IControlInterface,
)
from _mmdeviceapi import (
    IMMDeviceEnumerator,
    IMMNotificationClient,
    MMDeviceEnumerator,
    IMMEndpoint,
    PIMMEndpoint
)
from _audiopolicy import PIAudioSessionManager2
from _devicetopologyapi import (
    IControlInterface,
    PIDeviceTopology,
    IDeviceTopology,
    IPart,
    PIAudioBass,
    PIKsJackDescription2,
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
from _constant import (
    PKEY_Device_FriendlyName,
    PKEY_Device_DeviceDesc,
    PKEY_DeviceInterface_FriendlyName,
    STGM_READ,
    DEVICE_STATE_MASK_ALL,
    KSNODETYPE_DAC,
    KSNODETYPE_ADC,
    KSNODETYPE_SRC,
    KSNODETYPE_SUPERMIX,
    KSNODETYPE_MUX,
    KSNODETYPE_DEMUX,
    KSNODETYPE_SUM,
    KSNODETYPE_MUTE,
    KSNODETYPE_VOLUME,
    KSNODETYPE_TONE,
    KSNODETYPE_EQUALIZER,
    KSNODETYPE_AGC,
    KSNODETYPE_NOISE_SUPPRESS,
    KSNODETYPE_DELAY,
    KSNODETYPE_LOUDNESS,
    KSNODETYPE_PROLOGIC_DECODER,
    KSNODETYPE_STEREO_WIDE,
    KSNODETYPE_REVERB,
    KSNODETYPE_CHORUS,
    KSNODETYPE_3D_EFFECTS,
    KSNODETYPE_PARAMETRIC_EQUALIZER,
    KSNODETYPE_UPDOWN_MIX,
    KSNODETYPE_DYN_RANGE_COMPRESSOR,
    KSNODETYPE_DEV_SPECIFIC,
    KSNODETYPE_PROLOGIC_ENCODER,
    KSNODETYPE_PEAKMETER,
    # KSNODETYPE_SURROUND_ENCODER,
    KSNODETYPE_ACOUSTIC_ECHO_CANCEL,
    KSNODETYPE_MICROPHONE_ARRAY_PROCESSOR
)

from _enum import (
    EDataFlow,
    ERole,
    EndpointConnectorType,
    EndpointFormFactor,
    ConnectorType,
    AudioDeviceState,
)

ENDPOINT_CONNECTOR_TYPE = {
    EndpointConnectorType.eHostProcessConnector: 'HostProcessConnector',
    EndpointConnectorType.eOffloadConnector: 'OffloadConnector',
    EndpointConnectorType.eLoopbackConnector: 'LoopbackConnector',
    EndpointConnectorType.eKeywordDetectorConnector: 'KeywordDetectorConnector',
    EndpointConnectorType.eConnectorCount: 'ConnectorCount',
}

ENDPOINT_FORM_FACTOR = {
    EndpointFormFactor.RemoteNetworkDevice: 'RemoteNetworkDevice',
    EndpointFormFactor.Speakers: 'Speakers',
    EndpointFormFactor.LineLevel: 'LineLevel',
    EndpointFormFactor.Headphones: 'Headphones',
    EndpointFormFactor.Microphone: 'Microphone',
    EndpointFormFactor.Headset: 'Headset',
    EndpointFormFactor.Handset: 'Handset',
    EndpointFormFactor.UnknownDigitalPassthrough: 'UnknownDigitalPassthrough',
    EndpointFormFactor.SPDIF: 'SPDIF',
    EndpointFormFactor.DigitalAudioDisplayDevice: 'DigitalAudioDisplayDevice',
    EndpointFormFactor.UnknownFormFactor: 'UnknownFormFactor',
    EndpointFormFactor.EndpointFormFactor_enum_count: 'EndpointFormFactor_enum_count'
}

CONNECTOR_TYPE = {
    ConnectorType.Unknown_Connector: 'UnknownConnector',
    ConnectorType.Physical_Internal: 'PhysicalInternal',
    ConnectorType.Physical_External: 'PhysicalExternal',
    ConnectorType.Software_IO: 'SoftwareIO',
    ConnectorType.Software_Fixed: 'SoftwareFixed',
    ConnectorType.Network: 'Network'
}

AUDIO_DEVICE_STATE = {
    AudioDeviceState.Active: 'Active',
    AudioDeviceState.Disabled: 'Disabled',
    AudioDeviceState.NotPresent: 'NotPresent',
    AudioDeviceState.Unplugged: 'Unplugged'
}

E_DATA_FLOW = {
    EDataFlow.eRender: 'Render',
    EDataFlow.eCapture: 'Capture',
    EDataFlow.eAll: 'All',
    EDataFlow.EDataFlow_enum_count: 'EDataFlow_enum_count'
}

E_ROLE = {
    ERole.eConsole: 'Console',
    ERole.eMultimedia: 'Multimedia',
    ERole.eCommunications: 'Communications',
    ERole.ERole_enum_count: 'ERole_enum_count'
}


class _Value(object):
    def __init__(self, value):
        self.value = value

    def __int__(self):
        return self.value

    def __str__(self):
        raise NotImplementedError

    def __unicode__(self):
        return unicode(self.__str__())


class _EDataFlow(_Value):
    def __str__(self):
        return E_DATA_FLOW[self.value]


class _ERole(_Value):
    def __str__(self):
        return E_ROLE[self.value]


class _AudioDeviceState(_Value):
    def __str__(self):
        return AUDIO_DEVICE_STATE[self.value]


class _ConnectorType(_Value):
    def __str__(self):
        return CONNECTOR_TYPE[self.value]


class _EndpointConnectorType(_Value):
    def __str__(self):
        return ENDPOINT_CONNECTOR_TYPE[self.value]


class _EndpointFormFactor(_Value):
    def __str__(self):
        return ENDPOINT_FORM_FACTOR[self.value]


class _DeviceTopology(comtypes.COMObject):
    _com_interfaces_ = [IDeviceTopology]


class _NotificationClient(comtypes.COMObject):
    _com_interfaces_ = [IMMNotificationClient]

    def __init__(self, device_enum, callbacks):
        self.__device_enum = device_enum
        self.__callbacks = callbacks
        comtypes.COMObject.__init__(self)

    def OnDeviceStateChanged(self, pwstrDeviceId, dwNewState):
        device = self.__device_enum.GetDevice(pwstrDeviceId)
        state = _AudioDeviceState(dwNewState)
        for callback in self.__callbacks['state']:
            callback(device, state)

        print 'OnDeviceStateChanged', 'pwstrDeviceId:', pwstrDeviceId, 'dwNewState:', dwNewState

    def OnDeviceAdded(self, pwstrDeviceId):
        device = self.__device_enum.GetDevice(pwstrDeviceId)
        for callback in self.__callbacks['add']:
            callback(device)

        print 'OnDeviceAdded', 'pwstrDeviceId:', pwstrDeviceId

    def OnDeviceRemoved(self, pwstrDeviceId):
        device = self.__device_enum.GetDevice(pwstrDeviceId)
        for callback in self.__callbacks['remove']:
            callback(device)

        print 'OnDeviceRemoved', 'pwstrDeviceId:', pwstrDeviceId

    def OnDefaultDeviceChanged(self, flow, role, pwstrDefaultDeviceId):
        device = self.__device_enum.GetDevice(pwstrDefaultDeviceId)
        for callback in self.__callbacks['default']:
            callback(device)

        print 'OnDefaultDeviceChanged', 'flow:', flow, 'role:', role, 'pwstrDefaultDeviceId:', pwstrDefaultDeviceId

    def OnPropertyValueChanged(self, pwstrDeviceId, key):
        device = self.__device_enum.GetDevice(pwstrDeviceId)
        for callback in self.__callbacks['property']:
            callback(device, key)
        print 'OnPropertyValueChanged', 'pwstrDeviceId:', pwstrDeviceId, 'key:', key


class AudioPart(object):
    def __init__(self, parent, device):
        self._parent = parent
        self._device = device

    @property
    def __subtype(self):
        part = self._parent.QueryInterface(IPart)
        return part.GetSubType()

    @property
    def __interfaces(self):
        part = self._parent.QueryInterface(IPart)

        for index in range(part.GetControlInterfaceCount()):
            yield part.GetControlInterface(index)

    @property
    def auto_gain_control(self):
        if self.__subtype == KSNODETYPE_AGC:
            try:
                part = self._parent.QueryInterface(IPart)

                control = comtypes.cast(
                    part.Activate(
                        comtypes.CLSCTX_INPROC_SERVER,
                        IID_IAudioAutoGainControl
                    ),
                    PIAudioAutoGainControl
                )
                print 'IAudioAutoGainControl:', control

                return None
            except comtypes.COMError:
                raise AttributeError

    @property
    def bass(self):
        if self.__subtype == KSNODETYPE_TONE:
            try:
                part = self._parent.QueryInterface(IPart)

                control = comtypes.cast(
                    part.Activate(
                        comtypes.CLSCTX_INPROC_SERVER,
                        IID_IAudioBass
                    ),
                    PIAudioBass
                )
                print 'IAudioBass:', control
                return None
            except comtypes.COMError:
                raise AttributeError

    @property
    def channel_config(self):
        if self.__subtype in (
            KSNODETYPE_3D_EFFECTS,
            KSNODETYPE_DAC,
            KSNODETYPE_PROLOGIC_DECODER,
            KSNODETYPE_VOLUME
        ):
            try:
                part = self._parent.QueryInterface(IPart)

                control = comtypes.cast(
                    part.Activate(
                        comtypes.CLSCTX_INPROC_SERVER,
                        IID_IAudioChannelConfig
                    ),
                    PIAudioChannelConfig
                )
                print 'IAudioChannelConfig:', control

                return None
            except comtypes.COMError:
                raise AttributeError

    @property
    def input(self):
        if self.__subtype == KSNODETYPE_MUX:
            try:
                part = self._parent.QueryInterface(IPart)

                control = comtypes.cast(
                    part.Activate(
                        comtypes.CLSCTX_INPROC_SERVER,
                        IID_IAudioInputSelector
                    ),
                    PIAudioInputSelector
                )
                print 'IAudioInputSelector:', control
                return None
            except comtypes.COMError:
                raise AttributeError

    @property
    def loudness(self):
        if self.__subtype == KSNODETYPE_LOUDNESS:
            try:
                part = self._parent.QueryInterface(IPart)
                control = comtypes.cast(
                    part.Activate(
                        comtypes.CLSCTX_INPROC_SERVER,
                        IID_IAudioLoudness
                    ),
                    PIAudioLoudness
                )

                print 'IAudioLoudness:', control
                return None
            except comtypes.COMError:
                raise AttributeError

    @property
    def midrange(self):
        if self.__subtype == KSNODETYPE_TONE:
            try:
                part = self._parent.QueryInterface(IPart)

                control = comtypes.cast(
                    part.Activate(
                        comtypes.CLSCTX_INPROC_SERVER,
                        IID_IAudioMidrange
                    ),
                    PIAudioMidrange
                )
                print 'IAudioMidrange:', control
                return None
            except comtypes.COMError:
                raise AttributeError

    @property
    def mute(self):
        if self.__subtype == KSNODETYPE_MUTE:
            try:
                part = self._parent.QueryInterface(IPart)

                control = comtypes.cast(
                    part.Activate(
                        comtypes.CLSCTX_INPROC_SERVER,
                        IID_IAudioMute
                    ),
                    PIAudioMute
                )
                print 'IAudioMute:', control

                return None
            except comtypes.COMError:
                raise AttributeError

    @property
    def output(self):
        if self.__subtype == KSNODETYPE_DEMUX:
            try:
                part = self._parent.QueryInterface(IPart)

                control = comtypes.cast(
                    part.Activate(
                        comtypes.CLSCTX_INPROC_SERVER,
                        IID_IAudioOutputSelector
                    ),
                    PIAudioOutputSelector
                )
                print 'IAudioOutputSelector:', control

                return None
            except comtypes.COMError:
                raise AttributeError

    @property
    def peak_meter(self):
        if self.__subtype == KSNODETYPE_PEAKMETER:
            try:
                part = self._parent.QueryInterface(IPart)

                control = comtypes.cast(
                    part.Activate(
                        comtypes.CLSCTX_INPROC_SERVER,
                        IID_IAudioPeakMeter
                    ),
                    PIAudioPeakMeter
                )
                print 'IAudioPeakMeter:', control
                return None
            except comtypes.COMError:
                raise AttributeError

    @property
    def treble(self):
        if self.__subtype == KSNODETYPE_TONE:
            try:
                part = self._parent.QueryInterface(IPart)

                control = comtypes.cast(
                    part.Activate(
                        comtypes.CLSCTX_INPROC_SERVER,
                        IID_IAudioTreble
                    ),
                    PIAudioTreble
                )
                print 'IAudioTreble:', control

                return None
            except comtypes.COMError:
                raise AttributeError

    @property
    def volume(self):
        if self.__subtype == KSNODETYPE_VOLUME:
            try:
                part = self._parent.QueryInterface(IPart)

                control = comtypes.cast(
                    part.Activate(
                        comtypes.CLSCTX_INPROC_SERVER,
                        IID_IAudioVolumeLevel
                    ),
                    PIAudioVolumeLevel
                )
                print 'IAudioVolumeLevel:', control

                return None
            except comtypes.COMError:
                raise AttributeError

    @property
    def format_support(self):
        for interface in self.__interfaces:
            if str(IID_IKsFormatSupport) == str(interface.GetIID()):
                part = self._parent.QueryInterface(IPart)

                control = comtypes.cast(
                    part.Activate(
                        comtypes.CLSCTX_INPROC_SERVER,
                        IID_IKsFormatSupport
                    ),
                    PIKsFormatSupport
                )
                print 'IKsFormatSupport:', control

                return None


class AudioDeviceSubunit(AudioPart):
    pass


class AudioDeviceConnection(AudioPart):

    @property
    def connected_to_device_id(self):
        return self._parent.GetDeviceIdConnectedTo()

    @property
    def is_connected(self):
        return self._parent.IsConnected() == 1

    @property
    def connected_to(self):
        return AudioDeviceConnection(
            self._parent,
            self._parent.GetConnectedTo()
        )

    @connected_to.setter
    def connected_to(self, connection):
        self._parent.ConnectTo(connection.get_c_connector())

    @property
    def type(self):
        return _ConnectorType(self._parent.GetType().value)

    @property
    def data_flow(self):
        return _EDataFlow(self._parent.GetDataFlow().value)

    def get_c_connector(self):
        return self._parent

    def disconnect(self):
        self._parent.Disconnect()

    @property
    def _description(self):
        # for interface in self.__interfaces:
        #     if str(IID_IKsJackDescription) == str(interface.GetIID()):

        try:
            connected_to = self.connected_to.get_c_connector()
            part = connected_to.QueryInterface(IPart)

            control = comtypes.cast(
                part.Activate(
                    comtypes.CLSCTX_INPROC_SERVER,
                    IID_IKsJackDescription2
                ),
                PIKsJackDescription2
            )

            return control
        except comtypes.COMError:
            pass

    @property
    def name(self):
        """Return an endpoint devices FriendlyName."""

        description = self._description

        if description is None:
            pStore = self._device.OpenPropertyStore(STGM_READ)
            try:
                return pStore.GetValue(PKEY_Device_FriendlyName)
            except comtypes.COMError:
                pass
        return description

    @property
    def description(self):

        description = self._description

        if description is None:

            pStore = self._device.OpenPropertyStore(STGM_READ)
            try:
                return pStore.GetValue(PKEY_Device_DeviceDesc)
            except comtypes.COMError:
                pass

        return description

CLSCTX_INPROC_SERVER = comtypes.CLSCTX_INPROC_SERVER


import traceback
import _ctypes


class EndpointVolume(object):
    def __init__(self, device_enum, endpoint_enum, volume_interface):
        self.__device_enum = device_enum
        self.__endpoint_enum = endpoint_enum
        self.__volume_interface = volume_interface


class AudioEndpoint(object):
    def __init__(self, device_enum, endpoint_enum, endpoint):
        self.__device_enum = device_enum
        self.__endpoint_enum = endpoint_enum
        self.__endpoint = endpoint

    @property
    def audio_meter(self):
        try:
            meter = comtypes.cast(
                self.__endpoint.Activate(
                    IID_IAudioMeterInformation,
                    CLSCTX_INPROC_SERVER
                ),
                PIAudioMeterInformation
            )
            print '__meter:', meter

            return None
        except _ctypes.COMError:
            pass

    @property
    def audio_client(self):
        try:
            audio_client = comtypes.cast(
                self.__endpoint.Activate(
                    IID_IAudioClient,
                    CLSCTX_INPROC_SERVER
                ),
                PIAudioClient
            )
            print '__audio_client:', audio_client

            return None
        except _ctypes.COMError:
            pass

    @property
    def session_manager(self):
        try:
            session_namager = comtypes.cast(
                self.__endpoint.Activate(
                    IID_IAudioSessionManager2,
                    CLSCTX_INPROC_SERVER
                ),
                PIAudioSessionManager2
            )
            print '__session_namager:', session_namager

            return None

        except _ctypes.COMError:
            pass

    @property
    def data_flow(self):
        try:
            endpoint = self.__endpoint.QueryInterface(IMMEndpoint)
            return _EDataFlow(endpoint.GetDataFlow())
        except _ctypes.COMError:
            pass


class AudioDevice(object):

    def __init__(self, id, device_enum, data_flow):

        self.__device_enum = device_enum
        self.__id = id
        self.__data_flow = data_flow

        # endpoint_volume = comtypes.POINTER(IAudioEndpointVolumeEx)()

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
                        self.__device_topology.GetConnector(j),
                        self.__device
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
        return _AudioDeviceState(self.__device.GetState())

    @property
    def connector_count(self):
        return self.__device_topology.GetConnectorCount()

    @property
    def device_id(self):
        return self.__device_topology.GetDeviceId()

    @property
    def name(self):
        """Return an endpoint devices FriendlyName."""
        pStore = self.__device.OpenPropertyStore(STGM_READ)
        try:
            return pStore.GetValue(PKEY_DeviceInterface_FriendlyName)
        except _ctypes.COMError:
            pass


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

    @property
    def devices(self):
        return list(self.__devices(EDataFlow.eAll))

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
            except _ctypes.COMError:
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
        return iter(self.devices)

    def __getitem__(self, item):
        if isinstance(item, (slice, int)):
            return list(self.devices)[item]

        for device in self.devices:
            if item in (device.name, device.device_id):
                return device
        raise AttributeError









if __name__ == '__main__':

    def show_vol(ev):
        voldb = ev.GetMasterVolumeLevel()
        volsc = ev.GetMasterVolumeLevelScalar()
        volst, nstep = ev.GetVolumeStepInfo()
        print('Master Volume (dB): %0.4f' % voldb)
        print('Master Volume (scalar): %0.4f' % volsc)
        print('Master Volume (step): %d / %d' % (volst, nstep))

    def test():

        ad = AudioDevices()
        for device in ad:
            print 'name:', device.name
            for connector in device.connectors:
                print 'connector name:', connector.name
                print 'connector description:', connector.description

            print
            print


        return
        ev = IAudioEndpointVolumeEx.get_default()
        vol = ev.GetMasterVolumeLevelScalar()
        vmin, vmax, vinc = ev.GetVolumeRange()
        print('Volume Range (min, max, step) (dB): '
              '%0.4f, %0.4f, %0.4f' % (vmin, vmax, vinc))
        show_vol(ev)
        try:
            show_vol(ev)
            print '\nIncrement the master volume'
            ev.VolumeStepUp()
            show_vol(ev)
            print '\nIncrement the master volume'
            ev.VolumeStepUp()
            show_vol(ev)

            print '\nIncrement the master volume'
            ev.VolumeStepUp()
            show_vol(ev)
            print '\nDecrement the master volume twice'
            ev.VolumeStepDown()
            ev.VolumeStepDown()
            show_vol(ev)
            print '\nSet the master volume to 0.75 scalar'
            ev.SetMasterVolumeLevelScalar(0.75)
            show_vol(ev)
            print '\nSet the master volume to 0.25 scalar'
            ev.SetMasterVolumeLevelScalar(0.25)
            show_vol(ev)

            num_chan = ev.GetChannelCount()
            print '\nNumber of Channels: %d' % num_chan
            for i in range(num_chan):
                min_vol, max_vol, step_vol = ev.GetVolumeRangeChannel(i)
                print 'Channel: %d, Min Level: %ddB, Max Level: %ddB, Step: %ddB' % (
                    i, min_vol, max_vol, step_vol
                )

            original_levels = []
            for i in range(num_chan):
                chan_level_db = ev.GetChannelVolumeLevel(i)
                chan_level = ev.GetChannelVolumeLevelScalar(i)
                print 'Channel: %d, Scalar Level: %d, dB Level: %ddB' % (i, (chan_level * 100), chan_level_db)
                original_levels += [chan_level]

            ev.SetMasterVolumeLevelScalar(vol)

            import time

            for chan in range(num_chan):
                for level in range(101):

                    if level and not level % 5:
                        chan_level_db = ev.GetChannelVolumeLevel(chan)
                        chan_level = ev.GetChannelVolumeLevelScalar(chan)
                        print 'Channel: %d, Scalar Level: %d, dB Level: %ddB' % (chan, (chan_level * 100), chan_level_db)
                    time.sleep(0.05)
                    ev.SetChannelVolumeLevelScalar(chan, float(level / 100.0))
            time.sleep(0.5)
            for chan, level in enumerate(original_levels):
                ev.SetChannelVolumeLevelScalar(chan, level)

        finally:
            ev.SetMasterVolumeLevelScalar(vol)

    comtypes.CoInitialize()
    try:
        test()
        while True:
            pass
    finally:
        comtypes.CoUninitialize()
