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
from . import utils
from .mmdeviceapi import IMMDevice
from .audioclient import PISimpleAudioVolume, ISimpleAudioVolume
from .constant import S_OK
from .audiosessiontypes import (
    AudioSessionState,
    PAudioSessionState
)
from .signal import (
    tw,
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

IID_IAudioSessionEvents = IID(
    '{24918ACC-64B3-37C1-8CA9-74A66E9957A8}'
)
IID_IAudioSessionControl = IID(
    '{F4B1A599-7266-4319-A8CA-E70ACB11E8CD}'
)
IID_IAudioSessionControl2 = IID(
    '{BFB7FF88-7239-4FC9-8FA2-07C950BE9C6D}'
)
IID_IAudioSessionManager = IID(
    '{BFA971F1-4d5e-40bb-935e-967039bfbee4}'
)
IID_IAudioSessionManager2 = IID(
    '{77aa99a0-1bd6-484f-8bc7-2c654c9a9b6f}'
)
IID_IAudioSessionNotification = IID(
    '{641DD20B-4D41-49CC-ABA3-174B9477BB08}'
)
IID_IAudioVolumeDuckNotification = IID(
    '{C3B284D4-6D39-4359-B3CF-B56DDB3BB39C}'
)
IID_IAudioSessionEnumerator = IID(
    '{E2F5BB11-0570-40CA-ACDD-3AA01277DEE8}'
)


class AudioSessionDisconnectReason(ENUM):
    DisconnectReasonDeviceRemoval = 0
    DisconnectReasonServerShutdown = 1
    DisconnectReasonFormatChanged = 2
    DisconnectReasonSessionLogoff = 3
    DisconnectReasonSessionDisconnected = 4
    DisconnectReasonExclusiveModeOverride = 5


PAudioSessionDisconnectReason = POINTER(AudioSessionDisconnectReason)


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


# noinspection PyTypeChecker
PIAudioSessionEvents = POINTER(_IAudioSessionEvents)


class IAudioSessionEvents(comtypes.COMObject):
    _com_interfaces_ = [_IAudioSessionEvents]

    def __init__(self, session, endpoint):
        self.__session = session
        self.__endpoint = endpoint
        self.__session_name = session.name
        self.__session_icon = session.icon_path
        self.__last_volume_notification = (None, None)
        comtypes.COMObject.__init__(self)

    def OnDisplayNameChanged(self, NewDisplayName, _):
        NewDisplayName = utils.convert_to_string(NewDisplayName)

        ON_SESSION_NAME_CHANGED.signal(
            device=self.__endpoint.device,
            endpoint=self.__endpoint,
            session=self.__session,
            old_name=self.__session_name,
            new_name=NewDisplayName
        )

        self.__session_name = NewDisplayName

        return S_OK

    def OnIconPathChanged(self, NewIconPath, _):
        NewIconPath = utils.convert_to_string(NewIconPath)
        ON_SESSION_ICON_CHANGED.signal(
            device=self.__endpoint.device,
            endpoint=self.__endpoint,
            session=self.__session,
            old_icon=self.__session_icon,
            new_icon=NewIconPath
        )

        self.__session_icon = NewIconPath
        return S_OK

    def OnSimpleVolumeChanged(self, NewVolume, NewMute, _):
        if (NewVolume, NewMute) != self.__last_volume_notification:
            self.__last_volume_notification = (NewVolume, NewMute)

            endpoint_volume = self.__endpoint.volume.level
            new_volume = NewVolume * endpoint_volume

            ON_SESSION_VOLUME_CHANGED.signal(
                device=self.__endpoint.device,
                endpoint=self.__endpoint,
                session=self.__session,
                new_volume=new_volume,
                new_mute=NewMute
            )
        return S_OK

    def OnChannelVolumeChanged(self, _, NewChannelVolumeArray, ChangedChannel, __):
        vol = NewChannelVolumeArray[ChangedChannel]
        ON_SESSION_CHANNEL_VOLUME_CHANGED.signal(
            device=self.__endpoint.device,
            endpoint=self.__endpoint,
            session=self.__session,
            channel=ChangedChannel,
            new_volume=vol * 100.0
        )
        return S_OK

    def OnGroupingParamChanged(self, NewGroupingParam, _):
        ON_SESSION_GROUPING_CHANGED.signal(
            device=self.__endpoint.device,
            endpoint=self.__endpoint,
            session=self.__session,
            new_grouping_param=NewGroupingParam.contents
        )
        return S_OK

    def OnStateChanged(self, NewState, _):
        NewState = AudioSessionState.get(NewState)
        ON_SESSION_STATE_CHANGED.signal(
            device=self.__endpoint.device,
            endpoint=self.__endpoint,
            session=self.__session,
            new_state=NewState
        )
        return S_OK

    def OnSessionDisconnected(self, DisconnectReason, _):
        DisconnectReason = AudioSessionDisconnectReason.get(DisconnectReason)
        ON_SESSION_DISCONNECT.signal(
            device=self.__endpoint.device,
            endpoint=self.__endpoint,
            name=self.__session_name,
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
        self.__endpoint = None
        self.__session_events = None
        self.__session_manager = None
        self.__volume = None
        comtypes.IUnknown.__init__(self)

    def __call__(self, endpoint, session_manager):
        self.__endpoint = endpoint
        self.__session_manager = session_manager

        if self.__session_events is None:
            self.__session_events = IAudioSessionEvents(self, endpoint)
            # noinspection PyUnresolvedReferences
            self.RegisterAudioSessionNotification(self.__session_events)
        return self

    def __del__(self):
        if self.__session_events is not None:
            # noinspection PyUnresolvedReferences
            self.UnregisterAudioSessionNotification(self.__session_events)

    @property
    def state(self) -> ENUM_VALUE:
        """
        Get the session state
        """
        # noinspection PyUnresolvedReferences
        state = self.GetState().value
        return AudioSessionState.get(state)

    @property
    def session_manager(self):
        return self.__session_manager

    @property
    def endpoint(self) -> IMMDevice:
        """
        Endpoint the session belongs to
        """
        return self.__endpoint

    @property
    def icon_path(self) -> str:
        """
        Path to the icon ised for the session
        """
        # noinspection PyUnresolvedReferences
        data = self.GetIconPath()
        return utils.convert_to_string(data)

    @icon_path.setter
    def icon_path(self, value):
        raise NotImplementedError

    @property
    def name(self) -> str:
        """
        Name of the session
        """
        # noinspection PyUnresolvedReferences
        data = self.GetDisplayName()
        name = utils.convert_to_string(data)

        _CoTaskMemFree(data)

        if name == r'@%SystemRoot%\System32\AudioSrv.Dll,-202':
            name = 'System Sounds'

        name = name.strip()

        if not name:

            try:
                name = _get_process_name(self.process_id)
            except NotImplementedError:
                name = None

            if name is None:
                name = ''
                try:
                    id_ = self.id
                except AttributeError:
                    pass
                else:
                    id_ = id_.rsplit('%b', 1)[0]
                    id_ = id_.rsplit('\\', 1)[-1]
                    name = id_.split('.', 1)[0]
        return name

    @name.setter
    def name(self, value):
        raise NotImplementedError

    @property
    def grouping_param(self) -> GUID:
        """
        Not sure how this works or what it does
        """
        # noinspection PyUnresolvedReferences
        guid = self.GetGroupingParam()
        return guid

    @grouping_param.setter
    def grouping_param(self, guid: GUID):
        # noinspection PyUnresolvedReferences
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
        if self.__volume is None:
            vol = self.QueryInterface(ISimpleAudioVolume)
            if vol:
                self.__volume = vol(self)

        return self.__volume


# noinspection PyTypeChecker
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


# noinspection PyTypeChecker
PIAudioSessionNotification = POINTER(_IAudioSessionNotification)


class IAudioSessionNotification(comtypes.COMObject):
    _com_interfaces_ = [_IAudioSessionNotification]

    def __init__(self, session_manager, endpoint):
        self.__session_manager = session_manager
        self.__endpoint = endpoint

        comtypes.COMObject.__init__(self)

    def OnSessionCreated(self, NewSession):
        def _do(ns):

            session_control2 = ns.QueryInterface(IAudioSessionControl2)
            if session_control2:
                id_ = session_control2.instance_id
                for session in self.__session_manager:
                    if session.instance_id == id_:
                        break

                else:
                    return

                ON_SESSION_CREATED.signal(
                    device=self.__endpoint.device,
                    endpoint=self.__endpoint,
                    session=session
                )

        tw.add(_do, NewSession)

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
        # noinspection PyUnresolvedReferences
        count = self.GetCount()

        for i in range(count):
            # noinspection PyTypeChecker
            session_control = POINTER(IAudioSessionControl)()

            # noinspection PyUnresolvedReferences
            self.GetSession(INT(i), ctypes.byref(session_control))
            session_control2 = session_control.QueryInterface(IAudioSessionControl2)
            if session_control2:
                session_control = session_control2

            yield session_control


# noinspection PyTypeChecker
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
    def id(self) -> str:
        """
        Session id
        """
        # noinspection PyUnresolvedReferences
        data = self.GetSessionIdentifier()
        id_ = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return id_

    @property
    def instance_id(self) -> str:
        """
        Unique session id
        """
        # noinspection PyUnresolvedReferences
        data = self.GetSessionInstanceIdentifier()
        id_ = utils.convert_to_string(data)
        _CoTaskMemFree(data)
        return id_

    @property
    def process_id(self) -> int:
        """
        Process id
        """
        # noinspection PyUnresolvedReferences
        id_ = self.GetProcessId()
        return id_

    @property
    def is_system_sounds(self) -> bool:
        """
        Is this session for creatig system sound effects
        """
        # noinspection PyUnresolvedReferences
        return bool(self.IsSystemSoundsSession())

    def enable_ducking(self, value: bool):
        """
        Turn on/off fade in and out when transitioning between media streams
        """
        # noinspection PyUnresolvedReferences
        self.SetDuckingPreferences(BOOL(value))

    def __init__(self):
        IAudioSessionControl.__init__(self)
        self._vol_duck = None

    def __call__(self, endpoint, session_manager):
        IAudioSessionControl.__call__(self, endpoint, session_manager)

        self._vol_duck = IAudioVolumeDuckNotification(endpoint)

        # noinspection PyUnresolvedReferences
        session_manager.RegisterDuckNotification(self.GetSessionInstanceIdentifier(), self._vol_duck)
        return self

    def __del__(self):
        IAudioSessionControl.__del__(self)

        if self._vol_duck is not None:
            self.session_manager.UnregisterDuckNotification(self._vol_duck)


# noinspection PyTypeChecker
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
            (['out'], POINTER(PIAudioSessionControl), 'SessionControl')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetSimpleAudioVolume',
            (['in'], LPCGUID, 'AudioSessionGuid'),
            (['in'], DWORD, 'CrossProcessSession'),
            (['out'], POINTER(PISimpleAudioVolume), 'AudioVolume')
        )

    )

    def __init__(self):
        self.__endpoint = None
        comtypes.IUnknown.__init__(self)

    def __call__(self, endpoint):
        self.__endpoint = endpoint
        return self

    def __iter__(self):
        pass


# noinspection PyTypeChecker
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


# noinspection PyTypeChecker
PIAudioVolumeDuckNotification = POINTER(_IAudioVolumeDuckNotification)


class IAudioVolumeDuckNotification(comtypes.COMObject):
    _com_interfaces_ = [_IAudioVolumeDuckNotification]

    def __init__(self, session):
        self._session = session
        comtypes.COMObject.__init__(self)

    def OnVolumeDuckNotification(self, _, countCommunicationSessions):
        ON_SESSION_VOLUME_DUCK.signal(
            device=self._session.endpoint.device,
            endpoint=self._session.endpoint,
            session=self._session,
            count_communication_sessions=countCommunicationSessions
        )

        return S_OK

    def OnVolumeUnduckNotification(self, _):
        ON_SESSION_VOLUME_UNDUCK.signal(
            device=self._session.endpoint.device,
            endpoint=self._session.endpoint,
            session=self._session
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
                ['out'],
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
        self.__session_notification = IAudioSessionNotification(self, endpoint)

        # noinspection PyUnresolvedReferences
        self.RegisterSessionNotification(self.__session_notification)

        for _ in self:
            pass

        return self

    def __del__(self):
        if self.__session_notification is not None:
            # noinspection PyUnresolvedReferences
            self.UnregisterSessionNotification(self.__session_notification)

        self.__sessions.clear()

    def __iter__(self):
        """
        Get sessions
        """
        # noinspection PyUnresolvedReferences
        session_enum = self.GetSessionEnumerator()

        if session_enum:
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


# noinspection PyTypeChecker
PIAudioSessionManager2 = POINTER(IAudioSessionManager2)


# DWORD GetProcessImageFileNameW(
#   [in]  HANDLE hProcess,
#   [out] LPWSTR lpImageFileName,
#   [in]  DWORD  nSize
# );

_kernel32 = ctypes.windll.Kernel32
_psapi = ctypes.windll.Psapi

_GetProcessImageFileNameW = _psapi.GetProcessImageFileNameW
_GetProcessImageFileNameW.rstype = DWORD

# HANDLE OpenProcess(
#   [in] DWORD dwDesiredAccess,
#   [in] BOOL  bInheritHandle,
#   [in] DWORD dwProcessId
# );
_OpenProcess = _kernel32.OpenProcess
_OpenProcess.restype = HANDLE

# BOOL CloseHandle(
#   [in] HANDLE hObject
# );

_CloseHandle = _kernel32.CloseHandle
_CloseHandle.restype = BOOL

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

import os


def _get_process_name(id_):
    hprocess = _OpenProcess(
        DWORD(PROCESS_QUERY_LIMITED_INFORMATION),
        BOOL(1),
        DWORD(id_)
    )

    buf = ctypes.create_unicode_buffer(255)

    _GetProcessImageFileNameW(
        HANDLE(hprocess),
        buf,
        DWORD(255)
    )

    _CloseHandle(HANDLE(hprocess))

    value = buf.value

    if value:
        return os.path.split(os.path.splitext(value)[0])[-1]
