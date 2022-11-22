from .audioenginebaseapo import *
from .endpointvolumeapi import *
from .data_types import *
from .guiddef import *
import ctypes
import comtypes


class AUDIO_SYSTEMEFFECT_STATE(ENUM):
    AUDIO_SYSTEMEFFECT_STATE_OFF = 0
    AUDIO_SYSTEMEFFECT_STATE_ON = AUDIO_SYSTEMEFFECT_STATE_OFF + 1


class AUDIO_SYSTEMEFFECT(ctypes.Structure):
    _fields_ = [
        ('id', GUID),
        ('canSetState', BOOL),
        ('state', AUDIO_SYSTEMEFFECT_STATE)
    ]


IID_IAudioSystemEffects3 = IID(
    '{C58B31CD-FC6A-4255-BC1F-AD29BB0A4A17}'
)


class IAudioSystemEffects3(IAudioSystemEffects2):
    _case_insensitive_ = False
    _iid_ = IID_IAudioSystemEffects3
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetControllableSystemEffectsList',
            (['out'], POINTER(POINTER(AUDIO_SYSTEMEFFECT)), 'effects'),
            (['out'], POINTER(UINT), 'numEffects'),
            (['in'], HANDLE, 'event'),
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetAudioSystemEffectState',
            (['in'], GUID, 'effectId'),
            (['in'], AUDIO_SYSTEMEFFECT_STATE, 'state')
        )
    )


class APOInitSystemEffects3(ctypes.Structure):
    _fields_ = [
        ('APOInit', APOInitBaseStruct),
        ('pAPOEndpointProperties', POINTER(IPropertyStore)),
        ('pServiceProvider',  POINTER(comtypes.IServiceProvider)),
        ('pDeviceCollection',  POINTER(IMMDeviceCollection)),
        ('nSoftwareIoDeviceInCollection', UINT),
        ('nSoftwareIoConnectorIndex', UINT),
        ('AudioProcessingMode', GUID),
        ('InitializeForDiscoveryOnly', BOOL)
    ]


IID_IAudioProcessingObjectRTQueueService = IID(
    '{ACD65E2F-955B-4B57-B9BF-AC297BB752C9}'
)


class IAudioProcessingObjectRTQueueService(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioProcessingObjectRTQueueService
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetRealTimeWorkQueue',
            (['out'], POINTER(DWORD), 'workQueueId'),
        ),
    )


SID_AudioProcessingObjectRTQueue = DEFINE_GUID(
    0x458c1a1f,
    0x6899,
    0x4c12,
    0x99,
    0xac,
    0xe2,
    0xe6,
    0xac,
    0x25,
    0x31,
    0x4
)


class APO_LOG_LEVEL(ENUM):
    APO_LOG_LEVEL_ALWAYS = 0
    APO_LOG_LEVEL_CRITICAL = 1
    APO_LOG_LEVEL_ERROR = 2
    APO_LOG_LEVEL_WARNING = 3
    APO_LOG_LEVEL_INFO = 4
    APO_LOG_LEVEL_VERBOSE = 5


IID_IAudioProcessingObjectLoggingService = IID(
    '{698f0107-1745-4708-95a5-d84478a62a65}'
)


class IAudioProcessingObjectLoggingService(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioProcessingObjectLoggingService
    _methods_ = (
        COMMETHOD(
            [],
            VOID,
            'ApoLog',
            (['in'], APO_LOG_LEVEL, 'level'),
            (['in'], LPCWSTR, 'format')
        ),
    )
    

SID_AudioProcessingObjectLoggingService = DEFINE_GUID(
    0x8b8008af,
    0x9f9,
    0x456e,
    0xa1,
    0x73,
    0xbd,
    0xb5,
    0x84,
    0x99,
    0xbc,
    0xe7
)


class APO_NOTIFICATION_TYPE(ENUM):
    APO_NOTIFICATION_TYPE_NONE = 0
    APO_NOTIFICATION_TYPE_ENDPOINT_VOLUME = 1
    APO_NOTIFICATION_TYPE_ENDPOINT_PROPERTY_CHANGE = 2
    APO_NOTIFICATION_TYPE_SYSTEM_EFFECTS_PROPERTY_CHANGE = 3


class AUDIO_ENDPOINT_VOLUME_CHANGE_NOTIFICATION(ctypes.Structure):
    _fields_ = [
        ('endpoint', POINTER(IMMDevice)),
        ('volume', PAUDIO_VOLUME_NOTIFICATION_DATA)
    ]


class AUDIO_ENDPOINT_PROPERTY_CHANGE_NOTIFICATION(ctypes.Structure):
    _fields_ = [
        ('endpoint', POINTER(IMMDevice)),
        ('propertyStore', POINTER(IPropertyStore)),
        ('propertyKey', PROPERTYKEY)
    ]


class AUDIO_SYSTEMEFFECTS_PROPERTY_CHANGE_NOTIFICATION(ctypes.Structure):
    _fields_ = [
        ('endpoint', POINTER(IMMDevice)),
        ('propertyStoreContext', GUID),
        ('propertyStoreType', AUDIO_SYSTEMEFFECTS_PROPERTYSTORE_TYPE),
        ('propertyStore', POINTER(IPropertyStore)),
        ('propertyKey', PROPERTYKEY)
    ]


class APO_NOTIFICATION(ctypes.Structure):
    class DUMMYUNIONNAME(ctypes.Union):
        _fields_ = [
            (
                'audioEndpointVolumeChange',
                AUDIO_ENDPOINT_VOLUME_CHANGE_NOTIFICATION
            ),
            (
                'audioEndpointPropertyChange',
                AUDIO_ENDPOINT_PROPERTY_CHANGE_NOTIFICATION
            ),
            (
                'audioSystemEffectsPropertyChange',
                AUDIO_SYSTEMEFFECTS_PROPERTY_CHANGE_NOTIFICATION
            ),
        ]
    
    _fields_ = [
        ('type', APO_NOTIFICATION_TYPE),
        ('DUMMYUNIONNAME', DUMMYUNIONNAME)
    ]
    
    _anonymous_ = ('DUMMYUNIONNAME',)


class AUDIO_ENDPOINT_VOLUME_APO_NOTIFICATION_DESCRIPTOR(ctypes.Structure):
    _fields_ = [
        ('device', POINTER(IMMDevice))
    ]


class AUDIO_ENDPOINT_PROPERTY_CHANGE_APO_NOTIFICATION_DESCRIPTOR(
    ctypes.Structure
):
    _fields_ = [
        ('device', POINTER(IMMDevice))
    ]


class AUDIO_SYSTEMEFFECTS_PROPERTY_CHANGE_APO_NOTIFICATION_DESCRIPTOR(
    ctypes.Structure
):
    _fields_ = [
        ('device', POINTER(IMMDevice)),
        ('propertyStoreContext', GUID),
    ]


class APO_NOTIFICATION_DESCRIPTOR(ctypes.Structure):
    class DUMMYUNIONNAME(ctypes.Union):
        _fields_ = [ 
            (
                'audioEndpointVolume',
                AUDIO_ENDPOINT_VOLUME_APO_NOTIFICATION_DESCRIPTOR
            ),
            (
                'audioEndpointPropertyChange',
                AUDIO_ENDPOINT_PROPERTY_CHANGE_APO_NOTIFICATION_DESCRIPTOR
            ),
            (
                'audioSystemEffectsPropertyChange',
                AUDIO_SYSTEMEFFECTS_PROPERTY_CHANGE_APO_NOTIFICATION_DESCRIPTOR
            )
        ]
        
    _fields_ = [
        ('type', APO_NOTIFICATION_TYPE),
        ('DUMMYUNIONNAME', DUMMYUNIONNAME)
    ]
    
    _anonymous_ = ('DUMMYUNIONNAME',)


IID_IAudioProcessingObjectNotifications = IID(
    '{56B0C76F-02FD-4B21-A52E-9F8219FC86E4}'
)


class IAudioProcessingObjectNotifications(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioProcessingObjectNotifications
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetApoNotificationRegistrationInfo',
            (
                ['out'],
                POINTER(POINTER(APO_NOTIFICATION_DESCRIPTOR)),
                'apoNotifications'
            ),
            (['out'], POINTER(DWORD), 'count')
        ),
        COMMETHOD(
            [],
            VOID,
            'HandleNotification',
            (['in'], POINTER(APO_NOTIFICATION), 'apoNotification')
        )
    )
