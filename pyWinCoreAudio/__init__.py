import comtypes as _comtypes
import weakref as _weakref

from .signal import (
    ON_DEVICE_ADDED,
    ON_DEVICE_REMOVED,
    ON_DEVICE_PROPERTY_CHANGED,
    ON_DEVICE_STATE_CHANGED,
    ON_ENDPOINT_VOLUME_CHANGED,
    ON_ENDPOINT_DEFAULT_CHANGED,
    ON_SESSION_VOLUME_DUCK,
    ON_SESSION_VOLUME_UNDUCK,
    ON_SESSION_CREATED,
    ON_SESSION_NAME_CHANGED,
    ON_SESSION_GROUPING_CHANGED,
    ON_SESSION_ICON_CHANGED,
    ON_SESSION_DISCONNECT,
    ON_SESSION_VOLUME_CHANGED,
    ON_SESSION_STATE_CHANGED,
    ON_SESSION_CHANNEL_VOLUME_CHANGED,
    ON_PART_CHANGE
)

_device_enumerator = None


def devices(message=True):
    if message:
        print(
            '****************** IMPORTANT ******************\n'
            'DO NOT hold a reference to any of the objects\n'
            'that are produced from this library. If you do\n'
            'there is a high probability of getting a memory\n'
            'leak.\n\n'
            'Make sure you call pyWinCoreAudio.stop() when\n'
            'this library is no longer going to be used.\n\n'
            'The object returned from pyWinCoreAudio.devices\n'
            'is a weak reference and this is the only object\n'
            'from this library that you are allowed to hold\n'
            'a reference to. The object returned is a \n'
            'callable and must be called each and every\n'
            'time before using it. DO NOT hold a reference\n'
            'to the object that is returned.\n\n'
            'The above is really important so that the\n'
            'library can function properly and no errors\n'
            'will take place.'
            '***********************************************\n'
        )

    global _device_enumerator

    if _device_enumerator is None:
        _comtypes.CoInitialize()
        from .mmdeviceapi import IMMDeviceEnumerator

        _device_enumerator = IMMDeviceEnumerator()

    return _weakref.ref(_device_enumerator)


def stop():
    global _device_enumerator

    if _device_enumerator is not None:
        _device_enumerator = None
        _comtypes.CoUninitialize()
