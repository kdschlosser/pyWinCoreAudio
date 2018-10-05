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


import comtypes
import ctypes
import threading
from utils import run_in_thread, get_icon
from pyWinAPI.guiddef_h import GUID
from pyWinAPI.winerror_h import S_OK

from pyWinAPI.audiopolicy_h import (
    IAudioSessionEvents,
    IAudioSessionNotification,
    IAudioSessionControl2,
    IAudioSessionManager,
    IAudioSessionManager2,
    IAudioSessionEnumerator,
    AudioSessionDisconnectReason,
    AudioSessionState,
    IID_IAudioSessionManager,
    IID_IAudioSessionManager2
)

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

        run_in_thread(do)

        return S_OK


class AudioSessionManager(object):
    def __init__(self, endpoint):
        self.__endpoint = endpoint
        self.__sessions = {}
        self.__update_lock = threading.Lock()

        try:
            self.__session_manager = endpoint.activate(
                IID_IAudioSessionManager2,
                ctypes.POINTER(IAudioSessionManager2)
            )
        except comtypes.COMError:
            try:
                self.__session_manager = endpoint.activate(
                    IID_IAudioSessionManager,
                    ctypes.POINTER(IAudioSessionManager)
                )
            except comtypes.COMError:
                raise NotImplementedError

    def get_simple_volume(self, guid):
        return self.__session_manager.GetSimpleAudioVolume(
            ctypes.byref(guid), 0
        )

    def get_session_control(self, id):
        session_enum = self.__session_manager.GetSessionEnumerator()
        for i in range(session_enum.GetCount()):
            session = session_enum.GetSession(i)
            session_control = session.QueryInterface(IAudioSessionControl2)
            if session_control.GetSessionIdentifier() == id:
                return session_control

    def update_sessions(self, callback=None):

        self.__update_lock.acquire()

        event = threading.Event()
        event.wait(0.1)

        try:
            session_enum = self.__session_manager.GetSessionEnumerator()
            sessions = {}

            for i in range(session_enum.GetCount()):
                session = session_enum.GetSession(i)
                session_control = session.QueryInterface(IAudioSessionControl2)
                id = session_control.GetSessionIdentifier()

                if id in self.__sessions:
                    sessions[id] = self.__sessions[id]
                    del self.__sessions[id]

                else:
                    sessions[id] = AudioSession(self, id)
                    if callback is not None:
                        callback(sessions[id])

            for session in self.__sessions.values():
                session.release()

            self.__sessions.clear()

            for key, value in sessions.items():
                self.__sessions[key] = value
        except comtypes.COMError:
            pass

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

            for i in range(session_enum.GetCount()):
                session = session_enum.GetSession(i)
                session_control = session.QueryInterface(IAudioSessionControl2)
                id = session_control.GetSessionIdentifier()

                if id not in self.__sessions:
                    self.__sessions[id] = AudioSession(self, id)

                yield self.__sessions[id]
        except comtypes.COMError:
            pass


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

        run_in_thread(do)

        return S_OK

    def OnDisplayNameChanged(self, NewDisplayName, _):
        if NewDisplayName == '@%SystemRoot%\System32\AudioSrv.Dll,-202':
            NewDisplayName = 'System Sounds'

        def do():
            self.__callback.session_name_changed(
                self.__session,
                NewDisplayName,
            )

        run_in_thread(do)

        return S_OK

    def OnGroupingParamChanged(self, NewGroupingParam, _):

        def do():
            self.__callback.session_grouping_changed(
                self.__session,
                NewGroupingParam,
            )

        run_in_thread(do)

        return S_OK

    def OnIconPathChanged(self, NewIconPath, _):
        def do():
            self.__callback.session_icon_path_changed(
                self.__session,
                NewIconPath,
            )

        run_in_thread(do)

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

        run_in_thread(do)

        return S_OK

    def OnSimpleVolumeChanged(self, NewVolume, NewMute, _):

        def do():
            self.__callback.session_volume_changed(
                self.__session,
                NewVolume,
                bool(NewMute)
            )

        run_in_thread(do)

        return S_OK

    def OnStateChanged(self, NewState):

        def do(state):
            self.__callback.session_state_changed(
                self.__session,
                AUDIO_SESSION_STATE[state]
            )

        run_in_thread(do, NewState)

        return S_OK


class AudioSession(object):

    def __init__(self, session_manager, id):
        self.__session_manager = session_manager
        self.__id = id

    def release(self):
        pass

    @property
    def __session(self):
        return self.__session_manager.get_session_control(self.__id)

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
    def icon(self):
        return get_icon(self.__session.GetIconPath())

    @icon.setter
    def icon(self, icon_path):
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
        return self.__id

    @property
    def instance_id(self):
        try:
            return self.__session.GetSessionInstanceIdentifier()
        except comtypes.COMError:
            raise AttributeError

    @property
    def __simple_volume(self):
        return self.__session_manager.get_simple_volume(
            GUID('{' + self.id.rsplit('{')[-1])
        )

    @property
    def volume(self):
        return self.__simple_volume.GetMasterVolume() * 100.0

    @volume.setter
    def volume(self, level):
        self.__simple_volume.SetMasterVolume(float(level / 100.0), None)

    @property
    def mute(self):
        return self.__simple_volume.GetMute()

    @mute.setter
    def mute(self, mute):
        self.__simple_volume.SetMute(mute, None)

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
