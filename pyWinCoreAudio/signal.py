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
                if cb is None or cb == callback:
                    self._callbacks.remove(wcb)

    def signal(self, *args, **kwargs):
        for wcb in self._callbacks[:]:
            cb = wcb()
            if cb is None:
                self._callbacks.remove(wcb)
            else:
                _tw.add(cb, self, *args, **kwargs)


class SessionSignal(object):

    def __init__(self):
        self._callbacks = {}

    def register(self, callback, device=None, endpoint=None, session=None):
        if session is not None and endpoint is None:
            raise ValueError('You must specify an endpoint if you are specifying a session')

        if endpoint is not None and device is None:
            raise ValueError('You must specify a device if you are specifying an endpoint')

        if device is not None:
            device = weakref.ref(device)
        if endpoint is not None:
            endpoint = weakref.ref(endpoint)
        if session is not None:
            session = weakref.ref(session)

        callback = weakref.ref(callback)

        key = (device, endpoint, session)

        if key not in self._callbacks:
            self._callbacks[key] = []

        if callback not in self._callbacks[key]:
            self._callbacks[key].append(callback)

        return SignalRet(self, callback)

    def unregister(self, callback):

        for key, callbacks in list(self._callbacks.items()):
            for wcb in callbacks[:]:
                cb = wcb()
                if cb is None or cb == callback:
                    callbacks.remove(wcb)

                    if not callbacks:
                        del self._callbacks[key]

    def signal(self, device, endpoint, session, *args, **kwargs):

        for (wd, we, ws), callbacks in list(self._callbacks.items()):
            if wd is None:
                d = device
            else:
                d = wd()
            if we is None:
                e = endpoint
            else:
                e = we()
            if ws is None:
                s = session
            else:
                s = ws()

            if None in (d, e, s):
                del self._callbacks[(wd, we, ws)]
                continue

            if d == device and e == endpoint and s == session:
                for wcb in callbacks[:]:
                    cb = wcb()

                    if cb is None:
                        callbacks.remove(wcb)
                        if not callbacks:
                            del self._callbacks[(wd, we, ws)]
                            break

                        continue

                    _tw.add(cb, self, device, endpoint, session, *args, **kwargs)


class EndpointSignal(object):

    def __init__(self):
        self._callbacks = {}

    def register(self, callback, device=None, endpoint=None):
        if endpoint is not None and device is None:
            raise ValueError('You must specify a device if you are specifying an endpoint')

        if device is not None:
            device = weakref.ref(device)
        if endpoint is not None:
            endpoint = weakref.ref(endpoint)

        callback = weakref.ref(callback)

        key = (device, endpoint)

        if key not in self._callbacks:
            self._callbacks[key] = []

        if callback not in self._callbacks[key]:
            self._callbacks[key].append(callback)

        return SignalRet(self, callback)

    def unregister(self, callback):

        for key, callbacks in list(self._callbacks.items()):
            for wcb in callbacks[:]:
                cb = wcb()
                if cb is None or cb == callback:
                    callbacks.remove(wcb)

                    if not callbacks:
                        del self._callbacks[key]

    def signal(self, device, endpoint, *args, **kwargs):

        for (wd, we), callbacks in list(self._callbacks.items()):
            if wd is None:
                d = device
            else:
                d = wd()
            if we is None:
                e = endpoint
            else:
                e = we()

            if None in (d, e):
                del self._callbacks[(wd, we)]
                continue

            if d == device and e == endpoint:
                for wcb in callbacks[:]:
                    cb = wcb()

                    if cb is None:
                        callbacks.remove(wcb)
                        if not callbacks:
                            del self._callbacks[(wd, we)]
                            break

                        continue

                    _tw.add(cb, self, device, endpoint, *args, **kwargs)


class DeviceSignal(object):

    def __init__(self):
        self._callbacks = {}

    def register(self, callback, device=None):

        if device is not None:
            device = weakref.ref(device)

        callback = weakref.ref(callback)

        key = device

        if key not in self._callbacks:
            self._callbacks[key] = []

        if callback not in self._callbacks[key]:
            self._callbacks[key].append(callback)

        return SignalRet(self, callback)

    def unregister(self, callback):

        for key, callbacks in list(self._callbacks.items()):
            for wcb in callbacks[:]:
                cb = wcb()
                if cb is None or cb == callback:
                    callbacks.remove(wcb)

                    if not callbacks:
                        del self._callbacks[key]

    def signal(self, device, *args, **kwargs):

        for wd, callbacks in list(self._callbacks.items()):
            if wd is None:
                d = device
            else:
                d = wd()

            if d is None:
                del self._callbacks[wd]
                continue

            if d == device:
                for wcb in callbacks[:]:
                    cb = wcb()

                    if cb is None:
                        callbacks.remove(wcb)
                        if not callbacks:
                            del self._callbacks[wd]
                            break

                        continue

                    _tw.add(cb, self, device, *args, **kwargs)


class InterfaceSignal(object):

    def __init__(self):
        self._callbacks = {}

    def register(self, callback, device=None, endpoint=None, interface=None):
        if interface is not None and endpoint is None:
            raise ValueError('You must specify an endpoint if you are specifying an interface')

        if endpoint is not None and device is None:
            raise ValueError('You must specify a device if you are specifying an endpoint')

        if device is not None:
            device = weakref.ref(device)
        if endpoint is not None:
            endpoint = weakref.ref(endpoint)
        if interface is not None:
            session = weakref.ref(interface)

        callback = weakref.ref(callback)

        key = (device, endpoint, interface)

        if key not in self._callbacks:
            self._callbacks[key] = []

        if callback not in self._callbacks[key]:
            self._callbacks[key].append(callback)

        return SignalRet(self, callback)

    def unregister(self, callback):

        for key, callbacks in list(self._callbacks.items()):
            for wcb in callbacks[:]:
                cb = wcb()
                if cb is None or cb == callback:
                    callbacks.remove(wcb)

                    if not callbacks:
                        del self._callbacks[key]

    def signal(self, device, endpoint, interface, *args, **kwargs):

        for (wd, we, wi), callbacks in list(self._callbacks.items()):
            if wd is None:
                d = device
            else:
                d = wd()
            if we is None:
                e = endpoint
            else:
                e = we()
            if wi is None:
                i = interface
            else:
                i = wi()

            if None in (d, e, i):
                del self._callbacks[(wd, we, wi)]
                continue

            if d == device and e == endpoint and i == interface:
                for wcb in callbacks[:]:
                    cb = wcb()

                    if cb is None:
                        callbacks.remove(wcb)
                        if not callbacks:
                            del self._callbacks[(wd, we, wi)]
                            break

                        continue

                    _tw.add(cb, self, device, endpoint, interface, *args, **kwargs)


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
                        func(args, kwargs)
                    except:
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


_tw = SignalThreadWorker()

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
