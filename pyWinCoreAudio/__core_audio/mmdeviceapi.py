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
import ctypes
import comtypes
from comtypes import CoClass
from .. import utils
from .policyconfig import IPolicyConfigVista
from .audioclient import IAudioClient
from .audiopolicy import (
    IAudioSessionManager,
    IAudioSessionManager2,
)
from .endpointvolumeapi import (
    IAudioEndpointVolumeEx,
    IAudioEndpointVolume
)
from .devicetopologyapi import (
    IDeviceTopology,
    AudioSpeakers,
    IAudioInputSelector,
    IKsJackDescription,
    IKsJackDescription2,
    JackDescription,
    IKsJackSinkInformation,
    IAudioAutoGainControl,
    IAudioBass,
    IAudioChannelConfig,
    IAudioLoudness,
    IAudioMidrange,
    IAudioOutputSelector,
    IAudioTreble,
    IConnector,
    ISubunit
)
from .propertystore import (
    PROPERTYKEY,
    PIPropertyStore,
    PPROPVARIANT
)
from .enum_constants import (
    ERole,
    EDataFlow,
    PEDataFlow,
    EndpointFormFactor
)
from .constant import (
    S_OK,
    NodeTypeGUID,
    PKEY_DeviceInterface_FriendlyName,
    STGM_READ,
    STGM_WRITE,
    DEVICE_STATE_MASK_ALL,
    PKEY_Device_FriendlyName,
    PKEY_Device_DeviceDesc,
    PKEY_AudioEndpoint_FormFactor,
    PKEY_AudioEndpoint_JackSubType,
    PKEY_AudioEndpoint_FullRangeSpeakers,
    PKEY_AudioEndpoint_GUID,
    PKEY_AudioEndpoint_PhysicalSpeakers,
    PKEY_AudioEndpoint_Disable_SysFx,
    DEVICE_STATE_UNPLUGGED,
    DEVICE_STATE_NOTPRESENT,
    DEVICE_STATE_DISABLED,
    DEVICE_STATE_ACTIVE,
)
from .iid import (
    IID_IDeviceTopology,
    IID_IMMDevice,
    IID_IMMDeviceCollection,
    IID_IMMDeviceEnumerator,
    IID_IMMNotificationClient,
    IID_IMMEndpoint,
    IID_IMMDeviceActivator,
    CLSID_MMDeviceEnumerator,
    CLSID_PolicyConfigVistaClient,
    IID_MMDeviceAPILib,
    IID_IActivateAudioInterfaceAsyncOperation,
    IID_IActivateAudioInterfaceCompletionHandler
)

from ..signal import (
    ON_DEVICE_STATE_CHANGED,
    ON_DEVICE_ADDED,
    ON_DEVICE_REMOVED,
    ON_ENDPOINT_DEFAULT_CHANGED,
    ON_DEVICE_PROPERTY_CHANGED
)


CLSCTX_INPROC_SERVER = comtypes.CLSCTX_INPROC_SERVER
PIUnknown = POINTER(comtypes.IUnknown)
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


PIMMNotificationClient = POINTER(_IMMNotificationClient)


class IMMNotificationClient(comtypes.COMObject):
    _com_interfaces_ = [_IMMNotificationClient]

    def __init__(self, device_enum):
        self.__device_enum = device_enum
        comtypes.COMObject.__init__(self)

    def OnDeviceStateChanged(self, pwstrDeviceId, dwNewState):
        pwstrDeviceId = utils.convert_to_string(pwstrDeviceId)

        if dwNewState == DEVICE_STATE_UNPLUGGED:
            state = 'Unplugged'
        elif dwNewState == DEVICE_STATE_NOTPRESENT:
            state = 'Not Present'
        elif dwNewState == DEVICE_STATE_DISABLED:
            state = 'Disabled'
        elif dwNewState == DEVICE_STATE_ACTIVE:
            state = 'Active'
        else:
            state = 'Unknown'

        for device in self.__device_enum:
            if device.id == pwstrDeviceId:
                ON_DEVICE_STATE_CHANGED.signal(device, new_state=state)
                break

            for endpoint in device:
                if endpoint.id == pwstrDeviceId:
                    ON_DEVICE_STATE_CHANGED.signal(device, endpoint=endpoint, new_state=state)
                    break
            else:
                continue

            break

        return S_OK

    def OnDeviceAdded(self, pwstrDeviceId):
        pwstrDeviceId = utils.convert_to_string(pwstrDeviceId)
        for device in self.__device_enum:
            if device.id == pwstrDeviceId:
                ON_DEVICE_ADDED.signal(device)
                break

        return S_OK

    def OnDeviceRemoved(self, pwstrDeviceId):
        pwstrDeviceId = utils.convert_to_string(pwstrDeviceId)

        name = self.__device_enum.get_device_name(pwstrDeviceId)

        for _ in self.__device_enum:
            continue

        ON_DEVICE_REMOVED.signal(name=name)
        return S_OK

    def OnDefaultDeviceChanged(self, flow, role, _):
        flow = EDataFlow.get(flow)
        role = ERole.get(role)
        endpoint = self.__device_enum.default_audio_endpoint(role, flow)

        ON_ENDPOINT_DEFAULT_CHANGED.signal(endpoint.device, endpoint=endpoint)
        return S_OK

    def OnPropertyValueChanged(self, pwstrDeviceId, key):
        pwstrDeviceId = utils.convert_to_string(pwstrDeviceId)

        for device in self.__device_enum:
            if device.id == pwstrDeviceId:
                ON_DEVICE_PROPERTY_CHANGED.signal(device, key=key)
                break

            for endpoint in device:
                if endpoint.id == pwstrDeviceId:
                    ON_DEVICE_PROPERTY_CHANGED.signal(device, endpoint=endpoint, key=key)
                    break
            else:
                continue

            break

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


PIMMEndpoint = POINTER(IMMEndpoint)


class IMMDevice(comtypes.IUnknown):
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
    def device_name(self):
        return self.get_property(PKEY_DeviceInterface_FriendlyName)

    def get_property(self, key):
        pStore = self.OpenPropertyStore(STGM_READ)
        try:
            return pStore.GetValue(key)
        except comtypes.COMError:
            raise AttributeError

    def set_property(self, key, value):
        pStore = self.OpenPropertyStore(STGM_WRITE)
        try:
            return pStore.GetValue(key)
        except comtypes.COMError:
            raise AttributeError

    def activate(self, cls):
        try:
            return ctypes.cast(
                self.Activate(
                    cls._iid_,
                    CLSCTX_INPROC_SERVER,
                    None
                ),
                POINTER(cls)
            )

        except comtypes.COMError:
            pass

    @property
    def state(self):
        return self.GetState()

    @property
    def id(self):
        data = self.GetId()
        id_ = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return id_

    def QueryInterface(self, interface, iid=None):
        if iid is None:
            iid = interface._iid_

        try:
            self._IUnknown__com_QueryInterface(ctypes.byref(iid), ctypes.byref(interface))
            return True
        except comtypes.COMError:
            return False

    @property
    def audio_client(self):
        audio_client = POINTER(IAudioClient)(self)
        self.activate(audio_client)
        return audio_client

    @property
    def session_manager(self):
        return self.__session_manager

    @property
    def data_flow(self):
        endpoint = POINTER(IMMEndpoint)()
        if self.QueryInterface(endpoint):
            return EDataFlow.get(endpoint.GetDataFlow().value)

    @property
    def name(self):
        """Return an endpoint devices FriendlyName."""
        val = self.get_property(PKEY_Device_FriendlyName)
        return utils.convert_to_string(val)

    @property
    def description(self):
        val = self.get_property(PKEY_Device_DeviceDesc)
        return utils.convert_to_string(val)

    @property
    def form_factor(self):
        form_factor = self.get_property(PKEY_AudioEndpoint_FormFactor)
        return EndpointFormFactor.get(form_factor)

    @property
    def type(self):
        sub_type = self.get_property(PKEY_AudioEndpoint_JackSubType)
        sub_type_guid = str(sub_type)
        value = NodeTypeGUID.get(sub_type_guid)
        if value is None:
            return 'Unknown'
        return value.label

    @property
    def full_range_speakers(self):
        return AudioSpeakers(
            self.get_property(PKEY_AudioEndpoint_FullRangeSpeakers)
        )

    @property
    def guid(self):
        return self.get_property(PKEY_AudioEndpoint_GUID)

    @property
    def physical_speakers(self):
        return AudioSpeakers(
            self.get_property(PKEY_AudioEndpoint_PhysicalSpeakers)
        )

    @property
    def system_effects(self):
        return bool(
            self.get_property(PKEY_AudioEndpoint_Disable_SysFx)
        )

    @property
    def hdcp_capable(self):
        js = self.jack_sink_information
        if js is None:
            return False
        return js.hdcp_capable

    @property
    def ai_capable(self):
        js = self.jack_sink_information
        if js is None:
            return False

        return js.ai_capable

    @property
    def connector_type(self):
        connector = self.connector.connected_to
        return connector.type

    @property
    def connector_name(self):
        connector = self.connector.connected_to
        part = connector.part

        return part.name

    @property
    def connector_subtype(self):
        connector = self.connector.connected_to
        part = connector.part

        return part.sub_type

    @property
    def connector_location(self):
        jd = self.jack_descriptions
        if not jd:
            return 'Unknown'

        jd = jd[0]

        return jd.location

    @property
    def connector_style(self):
        js = self.jack_sink_information
        if js is not None:
            return js.connection_type

        jd = self.jack_descriptions
        if not jd:
            return 'Unknown'

        jd = jd[0]

        return jd.connection_type

    @property
    def presence_detection(self):
        jd = self.jack_descriptions
        if not jd:
            return False

        jd = jd[0]

        return jd.presence_detection

    @property
    def connector_color(self):
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
    def is_connected(self):
        jd = self.jack_descriptions
        if not jd:
            return False

        jd = jd[0]

        return jd.is_connected

    @property
    def jack_descriptions(self):
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
    def jack_sink_information(self):
        conn_to = self.connector.connected_to
        if conn_to is None:
            return None

        part = conn_to.part
        sink_information = part.activate(IKsJackSinkInformation)

        if sink_information:
            return sink_information.GetJackSinkInformation()

    @property
    def auto_gain_control(self):
        return self.__get_interface(IAudioAutoGainControl._iid_)

    @property
    def bass(self):
        return self.__get_interface(IAudioBass._iid_)

    @property
    def channel_config(self):
        return self.__get_interface(IAudioChannelConfig._iid_)

    @property
    def input(self):
        return self.__get_interface(IAudioInputSelector._iid_)

    @property
    def loudness(self):
        return self.__get_interface(IAudioLoudness._iid_)

    @property
    def midrange(self):
        return self.__get_interface(IAudioMidrange._iid_)

    @property
    def output(self):
        return self.__get_interface(IAudioOutputSelector._iid_)

    def __get_interface(self, iid):
        for subunit in self.subunits:
            part = subunit.part
            for interface in part:
                if interface._iid_ == iid:
                    return interface
        #
        # device = self.__device
        # conn_from = device.connectors[0]
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
    def treble(self):
        return self.__get_interface(IAudioTreble._iid_)

    @property
    def volume(self):
        return self.__volume

    def __call__(self, device, device_topology):
        self.__device = device
        self.__device_topology = device_topology(endpoint=self)

        session_manager = self.activate(IAudioSessionManager2)
        if not session_manager:
            session_manager = self.activate(IAudioSessionManager)

        if session_manager:
            session_manager = session_manager(endpoint=self)
        else:
            session_manager = None

        self.__session_manager = session_manager

        endpoint_volume = self.activate(IAudioEndpointVolumeEx)

        if not endpoint_volume:
            endpoint_volume = self.activate(IAudioEndpointVolume)

        if endpoint_volume:
            endpoint_volume = endpoint_volume(endpoint=self)
        else:
            endpoint_volume = None

        self.__volume = endpoint_volume

        return self

    @property
    def connector(self):
        connector = POINTER(IConnector)()
        self.__device_topology.GetConnector(0, ctypes.byref(connector))
        return connector(endpoint=self)

    @property
    def subunits(self):
        res = []
        pCount = self.__device_topology.GetSubunitCount()

        for i in range(pCount):
            subunit = POINTER(ISubunit)()
            self.__device_topology.GetSubunit(i, ctypes.byref(subunit))
            res.append(subunit(endpoint=self))

        return res

    @property
    def device(self):
        return self.__device

    def set_default(self, role):
        policy_config = comtypes.CoCreateInstance(
            CLSID_PolicyConfigVistaClient,
            IPolicyConfigVista,
            comtypes.CLSCTX_ALL
        )

        policy_config.SetDefaultEndpoint(self.id, ERole.get(role))

    @property
    def is_default(self):
        return IMMDeviceEnumerator.default_audio_endpoint(self.data_flow, self.data_flow) == self

    def __iter__(self):
        if self.__session_manager is not None:
            for session in self.__session_manager:
                yield session


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
            connector = POINTER(IConnector)()
            self.__device_topology.GetConnector(i, ctypes.byref(connector))
            res.append(connector(endpoint=self))

        return res

    @property
    def subunits(self):
        res = []
        pCount = self.__device_topology.GetSubunitCount()

        for i in range(pCount):
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


PIMMDevice = POINTER(IMMDevice)


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
        count = self.GetCount()
        for i in range(count):
            item = POINTER(IMMDevice)()
            self.Item(UINT(i), ctypes.byref(item))
            yield item


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
            (['in'], POINTER(PIMMDevice), 'ppDevices')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetDevice',
            (['in'], LPCWSTR, 'pwstrId'),
            (['in'], POINTER(PIMMDevice), 'ppDevices')
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
        endpoint_enum = POINTER(IMMDeviceCollection)()

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

            self.__device_enum = comtypes.CoCreateInstance(
                CLSID_MMDeviceEnumerator,
                _IMMDeviceEnumerator,
                comtypes.CLSCTX_ALL
            )

            self.__notification_client = IMMNotificationClient(self)
            self.__device_enum.RegisterEndpointNotificationCallback(self.__notification_client)

        else:
            self.__dict__.update(IMMDeviceEnumerator._instance.__dict__)

    def __del__(self):
        if self.__notification_client is not None:
            self.__device_enum.UnregisterEndpointNotificationCallback(self.__notification_client)
            self.__notification_client = None

    def __iter__(self):
        for device in self.__device_enum:
            yield device

    @classmethod
    def default_audio_endpoint(cls, data_flow, role):
        self = cls._instance
        endp = POINTER(IMMDevice)()

        self.GetDefaultAudioEndpoint(
            data_flow,
            role,
            ctypes.byref(endp)
        )

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
