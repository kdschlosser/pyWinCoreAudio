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

from .data_types import *
import weakref
import comtypes
from .. import utils
from .audioclient import PISimpleAudioVolume, ISimpleAudioVolume
from .constant import S_OK
from .enum_constants import (
    AudioSessionState,
    PAudioSessionState,
    AudioSessionDisconnectReason
)
from .iid import (
    IID_IAudioSessionEvents,
    IID_IAudioSessionControl,
    IID_IAudioSessionControl2,
    IID_IAudioSessionManager,
    IID_IAudioSessionManager2,
    IID_IAudioSessionNotification,
    IID_IAudioVolumeDuckNotification,
    IID_IAudioSessionEnumerator,
)
from ..signal import (
    ON_SESSION_VOLUME_DUCK,
    ON_SESSION_VOLUME_UNDUCK,
    ON_SESSION_CREATED,
    ON_SESSION_NAME_CHANGED,
    ON_SESSION_CHANNEL_VOLUME_CHANGED,
    ON_SESSION_GROUPING_CHANGED,
    ON_SESSION_ICON_CHANGED,
    ON_SESSION_DISCONNECT,
    ON_SESSION_VOLUME_CHANGED,
    ON_SESSION_STATE_CHANGED
)

_CoTaskMemFree = ctypes.windll.ole32.CoTaskMemFree


class _IAudioSessionEvents(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioSessionEvents
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'OnDisplayNameChanged',
            (['in'], LPCWSTR, 'NewDisplayName'),
            (['in'], LPCGUID, 'EventContext'),
        ),
        COMMETHOD(
            [],
            HRESULT,
            'OnIconPathChanged',
            (['in'], LPWCHAR, 'NewIconPath'),
            (['in'], LPCGUID, 'EventContext'),
        ),
        COMMETHOD(
            [],
            HRESULT,
            'OnSimpleVolumeChanged',
            (['in'], FLOAT, 'NewVolume'),
            (['in'], BOOL, 'NewMute'),
            (['in'], LPCGUID, 'EventContext'),
        ),
        COMMETHOD(
            [],
            HRESULT,
            'OnChannelVolumeChanged',
            (['in'], DWORD, 'ChannelCount'),
            (['in'], (FLOAT * 8), 'NewChannelVolumeArray'),
            (['in'], DWORD, 'ChangedChannel'),
            (['in'], LPCGUID, 'EventContext'),
        ),
        COMMETHOD(
            [],
            HRESULT,
            'OnGroupingParamChanged',
            (['in'], LPCGUID, 'NewGroupingParam'),
            (['in'], LPCGUID, 'EventContext'),
        ),
        COMMETHOD(
            [],
            HRESULT,
            'OnStateChanged',
            (['in'], AudioSessionState, 'NewState'),
        ),
        COMMETHOD(
            [],
            HRESULT,
            'OnSessionDisconnected',
            (['in'], AudioSessionDisconnectReason, 'DisconnectReason'),
        )
    )


PIAudioSessionEvents = POINTER(_IAudioSessionEvents)


class IAudioSessionEvents(comtypes.COMObject):
    _com_interfaces_ = [_IAudioSessionEvents]

    def __init__(self, session):
        self.__session = session
        comtypes.COMObject.__init__(self)

    def OnDisplayNameChanged(self, NewDisplayName, _):
        NewDisplayName = utils.convert_to_string(NewDisplayName)

        ON_SESSION_NAME_CHANGED.signal(
            self.__session.endpoint.device,
            self.__session.endpoint,
            self.__session,
            new_name=NewDisplayName
        )
        return S_OK

    def OnIconPathChanged(self, NewIconPath, _):
        NewIconPath = utils.convert_to_string(NewIconPath)
        ON_SESSION_ICON_CHANGED.signal(
            self.__session.endpoint.device,
            self.__session.endpoint,
            self.__session,
            new_icon=NewIconPath
        )
        return S_OK

    def OnSimpleVolumeChanged(self, NewVolume, NewMute, _):
        ON_SESSION_VOLUME_CHANGED.signal(
            self.__session.endpoint.device,
            self.__session.endpoint,
            self.__session,
            new_volume=NewVolume,
            new_mute=NewMute
        )
        return S_OK

    def OnChannelVolumeChanged(self, _, NewChannelVolumeArray, ChangedChannel, __):
        vol = NewChannelVolumeArray[ChangedChannel]
        ON_SESSION_CHANNEL_VOLUME_CHANGED.signal(
            self.__session.endpoint.device,
            self.__session.endpoint,
            self.__session,
            channel=ChangedChannel,
            new_volume=vol
        )
        return S_OK

    def OnGroupingParamChanged(self, NewGroupingParam, _):
        ON_SESSION_GROUPING_CHANGED.signal(
            self.__session.endpoint.device,
            self.__session.endpoint,
            self.__session,
            new_grouping_param=NewGroupingParam.contents
        )
        return S_OK

    def OnStateChanged(self, NewState, _):
        NewState = AudioSessionState.get(NewState)
        ON_SESSION_STATE_CHANGED.signal(
            self.__session.endpoint.device,
            self.__session.endpoint,
            self.__session,
            new_state=NewState
        )
        return S_OK

    def OnSessionDisconnected(self, DisconnectReason, _):
        DisconnectReason = AudioSessionDisconnectReason.get(DisconnectReason)
        ON_SESSION_DISCONNECT.signal(
            self.__session.endpoint.device,
            self.__session.endpoint,
            self.__session,
            reason=DisconnectReason
        )
        session_manager = self.__session.session_manager()
        session_manager.remove_session(self.__session)
        return S_OK


class IAudioSessionControl(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioSessionControl
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetState',
            (['out'], PAudioSessionState, 'pRetVal')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetDisplayName',
            (['out'], POINTER(LPWSTR), 'pRetVal')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetDisplayName',
            (['in'], LPCWSTR, 'Value'),
            (['in', 'unique'], LPCGUID, 'EventContext')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetIconPath',
            (['out'], POINTER(LPCWSTR), 'pRetVal')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetIconPath',
            (['in'], LPCWSTR, 'Value'),
            (['in'], LPCGUID, 'EventContext')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetGroupingParam',
            (['out'], POINTER(GUID), 'pRetVal')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetGroupingParam',
            (['in'], LPCGUID, 'Override'),
            (['in'], LPCGUID, 'EventContext')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'RegisterAudioSessionNotification',
            (['in'], PIAudioSessionEvents, 'NewNotifications')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'UnregisterAudioSessionNotification',
            (['in'], PIAudioSessionEvents, 'NewNotifications')
        )
    )

    def __init__(self):
        self._endpoint = None
        self._session_events = None
        comtypes.IUnknown.__init__(self)

    def __call__(self, endpoint, session_manager):
        self._endpoint = endpoint
        self._session_manager = session_manager
        self._session_events = IAudioSessionEvents(self)
        self.RegisterAudioSessionNotification(self._session_events)
        return self

    def __del__(self):
        if self._session_events is not None:
            self.UnregisterAudioSessionNotification(self._session_events)

    def QueryInterface(self, interface, iid=None):

        if iid is None:
            iid = interface._iid_
        try:
            self.__com_QueryInterface(ctypes.byref(iid), ctypes.byref(interface))
            return True
        except comtypes.COMError:
            return False

    @property
    def session_manager(self):
        return self._session_manager

    @property
    def endpoint(self):
        return self._endpoint

    @property
    def icon_path(self):
        data = self.GetIconPath()
        return utils.convert_to_string(data)

    @icon_path.setter
    def icon_path(self, value):
        raise NotImplementedError

    @property
    def name(self):
        data = self.GetDisplayName()
        return utils.convert_to_string(data)

    @name.setter
    def name(self, value):
        raise NotImplementedError

    @property
    def grouping_param(self):
        guid = self.GetGroupingParam()
        return guid

    @grouping_param.setter
    def grouping_param(self, guid):
        self.SetGroupingParam(ctypes.byref(guid), NULL)

    @property
    def id(self):
        raise NotImplementedError

    @property
    def instance_id(self):
        raise NotImplementedError

    @property
    def process_id(self):
        raise NotImplementedError

    @property
    def is_system_sounds(self):
        raise NotImplementedError

    def enable_ducking(self, value):
        raise NotImplementedError

    @property
    def volume(self):
        return self._session_manager.get_volume()


PIAudioSessionControl = POINTER(IAudioSessionControl)


class _IAudioSessionNotification(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioSessionNotification
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'OnSessionCreated',
            (['in'], PIAudioSessionControl, 'NewSession')
        ),
    )


PIAudioSessionNotification = POINTER(_IAudioSessionNotification)


class IAudioSessionNotification(comtypes.COMObject):
    _com_interfaces_ = [_IAudioSessionNotification]

    def __init__(self, endpoint):
        self.__endpoint = endpoint
        comtypes.COMObject.__init__(self)

    def OnSessionCreated(self, NewSession):
        session_manager = self.__endpoint.session_manager

        name = utils.convert_to_string(NewSession.GetDisplayName())

        for session in session_manager:
            if session.name == name:
                break
        else:
            return S_OK

        ON_SESSION_CREATED.signal(self.__endpoint.device, self.__endpoint, session)
        return S_OK


class IAudioSessionEnumerator(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioSessionEnumerator
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetCount',
            (['out'], LPINT, 'SessionCount')
        ),

        COMMETHOD(
            [],
            HRESULT,
            'GetSession',
            (['in'], INT, 'SessionCount'),
            (['in'], POINTER(PIAudioSessionControl), 'Session')
        )
    )

    def __iter__(self):
        count = self.GetCount()

        for i in range(count):
            session_control = POINTER(IAudioSessionControl)()

            self.GetSession(INT(i), ctypes.byref(session_control))
            session_control2 = POINTER(IAudioSessionControl2)()
            if session_control.QueryInterface(session_control2):
                session_control = session_control2

            yield session_control


PIAudioSessionEnumerator = POINTER(IAudioSessionEnumerator)


class IAudioSessionControl2(IAudioSessionControl):
    _case_insensitive_ = False
    _iid_ = IID_IAudioSessionControl2
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetSessionIdentifier',
            (['out'], POINTER(LPCWSTR), 'pRetVal')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetSessionInstanceIdentifier',
            (['out'], POINTER(LPCWSTR), 'pRetVal')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetProcessId',
            (['out'], POINTER(DWORD), 'pRetVal')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'IsSystemSoundsSession'
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetDuckingPreferences',
            (['in'], BOOL, 'optOut')
        )
    )

    @property
    def id(self):
        data = self.GetSessionIdentifier()
        id_ = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return id_

    @property
    def instance_id(self):
        data = self.GetSessionInstanceIdentifier()
        id_ = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return id_

    @property
    def process_id(self):
        id_ = self.GetProcessId()
        return id_

    @property
    def is_system_sounds(self):
        return bool(self.IsSystemSoundsSession())

    def enable_ducking(self, value):
        self.SetDuckingPreferences(BOOL(value))

    def __init__(self):
        IAudioSessionControl.__init__(self)
        self._vol_duck = None

    def __call__(self, endpoint, session_manager):
        IAudioSessionControl.__call__(self, endpoint, session_manager)

        self._vol_duck = IAudioVolumeDuckNotification(endpoint)

        session_manager.RegisterDuckNotification(self.GetSessionInstanceIdentifier(), self._vol_duck)
        return self

    def __del__(self):
        IAudioSessionControl.__del__(self)

        if self._vol_duck is not None:
            self.session_manager.UnregisterDuckNotification(self._vol_duck)


PIAudioSessionControl2 = POINTER(IAudioSessionControl2)


class IAudioSessionManager(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioSessionManager
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetAudioSessionControl',
            (['in'], LPCGUID, 'AudioSessionGuid'),
            (['in'], DWORD, 'StreamFlags'),
            (['in'], POINTER(PIAudioSessionControl), 'SessionControl')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetSimpleAudioVolume',
            (['in'], LPCGUID, 'AudioSessionGuid'),
            (['in'], DWORD, 'CrossProcessSession'),
            (['in'], POINTER(PISimpleAudioVolume), 'AudioVolume')
        )

    )

    def __init__(self):
        self.__endpoint = None
        comtypes.IUnknown.__init__(self)

    def __call__(self, endpoint):
        self.__endpoint = endpoint
        return self

    def get_session_control(self, guid):
        session_control = POINTER(IAudioSessionControl)()
        self.GetAudioSessionControl(ctypes.byref(guid), DWORD(0), ctypes.byref(session_control))
        return session_control(self.__endpoint, self)

    def get_volume(self, guid, is_cross_process):
        vol = POINTER(ISimpleAudioVolume)()
        self.GetSimpleAudioVolume(ctypes.byref(guid), DWORD(is_cross_process), ctypes.byref(vol))

        return vol

    def __iter__(self):
        pass


PIAudioSessionManager = POINTER(IAudioSessionManager)


class _IAudioVolumeDuckNotification(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioVolumeDuckNotification
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'OnVolumeDuckNotification',
            (['in'], LPCWSTR, 'sessionID'),
            (['in'], UINT32, 'countCommunicationSessions'),
        ),
        COMMETHOD(
            [],
            HRESULT,
            'OnVolumeUnduckNotification',
            (['in'], LPCWSTR, 'sessionID')
        )
    )


PIAudioVolumeDuckNotification = POINTER(_IAudioVolumeDuckNotification)


class IAudioVolumeDuckNotification(comtypes.COMObject):
    _com_interfaces_ = [_IAudioVolumeDuckNotification]

    def __init__(self, session):
        self._session = session
        comtypes.COMObject.__init__(self)

    def OnVolumeDuckNotification(self, _, countCommunicationSessions):
        ON_SESSION_VOLUME_DUCK.signal(
            self._session.endpoint.device,
            self._session.endpoint,
            self._session,
            count_communication_sessions=countCommunicationSessions
        )

        return S_OK

    def OnVolumeUnduckNotification(self, _):
        ON_SESSION_VOLUME_UNDUCK.signal(
            self._session.endpoint.device,
            self._session.endpoint,
            self._session
        )
        return S_OK


class IAudioSessionManager2(IAudioSessionManager):
    _case_insensitive_ = False
    _iid_ = IID_IAudioSessionManager2
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetSessionEnumerator',
            (
                ['in'],
                POINTER(PIAudioSessionEnumerator),
                'SessionList'
            )
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'RegisterSessionNotification',
            (
                PIAudioSessionNotification,
            )
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'UnregisterSessionNotification',
            (
                PIAudioSessionNotification,
            )
        ),
        COMMETHOD(
            [],
            HRESULT,
            'RegisterDuckNotification',
            (['in'], LPCWSTR, 'SessionID'),
            (['in'], PIAudioVolumeDuckNotification, 'duckNotification')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'UnregisterDuckNotification',
            (['in'], PIAudioVolumeDuckNotification, 'duckNotification')
        )
    )

    def __init__(self):
        IAudioSessionManager.__init__(self)
        self.__session_notification = None
        self.__endpoint = None
        self.__sessions = {}

    def __call__(self, endpoint):
        self.__endpoint = endpoint
        self.__session_notification = IAudioSessionNotification(endpoint)

        self.RegisterSessionNotification(self.__session_notification)
        return self

    def remove_session(self, id_):
        if id_ in self.__sessions:
            del self.__sessions[id_]

    def __del__(self):
        if self.__session_notification is not None:
            self.UnregisterSessionNotification(self.__session_notification)

        self.__sessions.clear()

    def __iter__(self):
        session_enum = POINTER(IAudioSessionEnumerator)()
        self.GetSessionEnumerator(session_enum)

        session_ids = []

        for session in session_enum:
            id_ = session.instance_id
            session_ids.append(id_)

            if id_ not in self.__sessions:
                self.__sessions[id_] = session(self.__endpoint, self)

        for id_ in list(self.__sessions.keys())[:]:
            if id_ not in session_ids:
                del self.__sessions[id_]

        for session in self.__sessions.values():
            yield session


PIAudioSessionManager2 = POINTER(IAudioSessionManager2)
