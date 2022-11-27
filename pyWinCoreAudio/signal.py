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

import threading
from collections import deque
import traceback
import weakref


class SignalRet(object):

    def __init__(self, signal, callback):
        self._signal = signal
        self._callback = callback

    def unregister(self):
        if self._callback is not None:
            self._signal.unregister(self._callback)
            self._callback = None

    def __del__(self):
        self.unregister()

    def __eq__(self, other):
        return self._signal == other

    def __ne__(self, other):
        return not self.__eq__(other)

    def __hash__(self):
        return hash(str(self))


class Signal(object):

    def __init__(self):
        self._callbacks = []

    def register(self, callback):
        callback = weakref.ref(callback)
        self._callbacks.append(callback)
        return SignalRet(self, callback)

    def unregister(self, callback):
        if callback in self._callbacks:
            self._callbacks.remove(callback)
        else:
            for wcb in self._callbacks[:]:
                cb = wcb()
                if cb is None or cb == callback or wcb == callback:
                    self._callbacks.remove(wcb)

    def signal(self, *args, **kwargs):
        for wcb in self._callbacks[:]:
            cb = wcb()
            if cb is None:
                self._callbacks.remove(wcb)
            else:
                tw.add(cb, signal=self, *args, **kwargs)


class SessionSignal(object):

    def __init__(self):
        self._callbacks = {}

    def register(self, callback, device=None, endpoint=None, session=None):


        if session is not None:
            if endpoint is None:
                endpoint = session.endpoint

            if session.endpoint != endpoint:
                raise ValueError(
                    "session endpoint does not match supplied endpoint"
                )

        if endpoint is not None:
            if device is None:
                device = endpoint.device

            if endpoint.device != device:
                raise ValueError('endpoint device does not match supplied device')

        if device is not None:
            device = device.id

        if endpoint is not None:
            endpoint = endpoint.id

        if session is not None:
            session = session.id

        callback = weakref.ref(callback)

        key = (device, endpoint, session)

        self._callbacks[key] = callback
        return SignalRet(self, callback)

    def unregister(self, callback):
        for key, wcb in list(self._callbacks.items())[:]:
            cb = wcb()
            if cb is None or cb == callback or wcb == callback:
                del self._callbacks[key]

    def signal(self, device, endpoint, session, *args, **kwargs):
        for (wd, we, ws), wcb in list(self._callbacks.items())[:]:
            if wd is None:
                d = device.id
            else:
                d = wd
            if we is None:
                e = endpoint.id
            else:
                e = we
            if ws is None:
                s = session.id
            else:
                s = ws

            if d == device.id and e == endpoint.id and s == session.id:
                cb = wcb()

                if cb is None:
                    del self._callbacks[(wd, we, ws)]
                else:
                    tw.add(cb, self, device, endpoint, session, *args, **kwargs)


class EndpointSignal(object):

    def __init__(self):
        self._callbacks = {}

    def register(self, callback, device=None, endpoint=None):
        if endpoint is not None:

            if device is None:
                device = endpoint.device

            if endpoint.device != device:
                raise ValueError(
                    'endpoint device does not match supplied device'
                )

        if device is not None:
            device = device.id
        if endpoint is not None:
            endpoint = endpoint.id

        callback = weakref.ref(callback)

        key = (device, endpoint)
        self._callbacks[key] = callback

        return SignalRet(self, callback)

    def unregister(self, callback):

        for key, wcb in list(self._callbacks.items())[:]:
            cb = wcb()
            if cb is None or cb == callback or wcb == callback:
                del self._callbacks[key]

    def signal(self, device, endpoint=None, *args, **kwargs):
        for (wd, we), wcb in list(self._callbacks.items())[:]:
            if wd is None:
                d = device.id
            else:
                d = wd
            if we is None:
                e = endpoint.id
            else:
                e = we

            if d == device.id and e == endpoint.id:
                cb = wcb()

                if cb is None:
                    del self._callbacks[(wd, we)]
                else:
                    tw.add(cb, self, device, endpoint, *args, **kwargs)


class DeviceSignal(object):

    def __init__(self):
        self._callbacks = {}

    def register(self, callback, device=None):

        if device is not None:
            device = device.id

        callback = weakref.ref(callback)

        key = device

        self._callbacks[key] = callback
        return SignalRet(self, callback)

    def unregister(self, callback):
        for key, wcb in list(self._callbacks.items())[:]:
            cb = wcb()
            if cb is None or cb == callback or wcb == callback:
                del self._callbacks[key]

    def signal(self, device, *args, **kwargs):
        for wd, wcb in list(self._callbacks.items())[:]:
            if wd is None:
                d = device.id
            else:
                d = wd

            if d == device.id:
                cb = wcb()

                if cb is None:
                    del self._callbacks[wd]
                else:
                    tw.add(cb, self, device, *args, **kwargs)


class InterfaceSignal(object):

    def __init__(self):
        self._callbacks = {}

    def register(self, callback, device=None, endpoint=None, interface=None):
        if interface is not None:
            if endpoint is None:
                endpoint = interface.endpoint

            if interface.endpoint != endpoint:
                raise ValueError(
                    'interface endpoint does not match supplied endpoint'
                )

        if endpoint is not None:
            if device is None:
                device = endpoint.device

            if endpoint.device != device:
                raise ValueError(
                    'endpoint device does not match supplied device'
                )

        if device is not None:
            device = device.id
        if endpoint is not None:
            endpoint = endpoint.id
        if interface is not None:
            interface = interface.global_id

        callback = weakref.ref(callback)

        key = (device, endpoint, interface)

        self._callbacks[key] = callback
        return SignalRet(self, callback)

    def unregister(self, callback):
        for key, wcb in list(self._callbacks.items())[:]:
            cb = wcb()
            if cb is None or cb == callback or wcb == callback:
                del self._callbacks[key]

    def signal(self, device, endpoint, interface, *args, **kwargs):
        for (wd, we, wi), wcb in list(self._callbacks.items()):
            if wd is None:
                d = device.id
            else:
                d = wd
            if we is None:
                e = endpoint.id
            else:
                e = we.id
            if wi is None:
                i = interface.global_id
            else:
                i = wi

            if d == device.id and e == endpoint.id and i == interface.global_id:
                cb = wcb()

                if cb is None:
                    del self._callbacks[(wd, we, wi)]
                else:
                    tw.add(cb, self, device, endpoint, interface, *args, **kwargs)


class SignalThreadWorker(object):

    def __init__(self):
        self._thread = None
        self._exit_event = threading.Event()
        self._queue_event = threading.Event()
        self._add_lock = threading.Lock()
        self._queue = deque()

    def _loop(self):

        while not self._exit_event.is_set():
            self._queue_event.wait(60)
            if self._queue_event.is_set():
                self._queue_event.clear()

                while True:
                    try:
                        func, args, kwargs = self._queue.popleft()
                    except IndexError:
                        break

                    try:
                        func(*args, **kwargs)
                    except:  # NOQA
                        traceback.print_exc()
            else:
                self._exit_event.set()

        self._thread = None
        self._exit_event.clear()

    def add(self, func, *args, **kwargs):
        with self._add_lock:
            self._queue.append((func, args, kwargs))
            self._queue_event.set()

            if self._thread is None:
                self._thread = threading.Thread(target=self._loop)
                self._thread.daemon = True
                self._thread.start()


tw = SignalThreadWorker()

ON_DEVICE_ADDED = Signal()
ON_DEVICE_REMOVED = Signal()
ON_DEVICE_PROPERTY_CHANGED = DeviceSignal()  # device
ON_DEVICE_STATE_CHANGED = DeviceSignal()  # device
ON_ENDPOINT_VOLUME_CHANGED = EndpointSignal()  # device, endpoint
ON_ENDPOINT_DEFAULT_CHANGED = EndpointSignal()  # device, endpoint

ON_SESSION_VOLUME_DUCK = SessionSignal()  # device, endpoint, session
ON_SESSION_VOLUME_UNDUCK = SessionSignal()  # device, endpoint, session
ON_SESSION_CREATED = EndpointSignal()  # device, endpoint
ON_SESSION_NAME_CHANGED = SessionSignal()  # device, endpoint, session
ON_SESSION_GROUPING_CHANGED = SessionSignal()  # device, endpoint, session
ON_SESSION_ICON_CHANGED = SessionSignal()  # device, endpoint, session
ON_SESSION_DISCONNECT = SessionSignal()  # device, endpoint, session
ON_SESSION_VOLUME_CHANGED = SessionSignal()  # device, endpoint, session
ON_SESSION_STATE_CHANGED = SessionSignal()  # device, endpoint, session
ON_SESSION_CHANNEL_VOLUME_CHANGED = SessionSignal()  # device, endpoint, session

ON_PART_CHANGE = InterfaceSignal()  # device, endpoint, interface
