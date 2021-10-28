
from .ks import *
from .data_types import *


STATIC_IID_IKsControl = (0x28F54685, 0x06FD, 0x11D2, 0xB2, 0x7A, 0x00, 0xA0, 0xC9, 0x22, 0x31, 0x96)
IID_IKsControl = DEFINE_GUIDEX(*STATIC_IID_IKsControl)


class IKsControl(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IKsControl
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'KsProperty',
            (['in'], PKSPROPERTY, 'Property'),
            (['in'], ULONG, 'PropertyLength'),
            (['in'], LPVOID, 'PropertyData'),
            (['in'], ULONG, 'DataLength'),
            (['in'], POINTER(ULONG), 'BytesReturned')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'KsMethod',
            (['in'], PKSMETHOD, 'Method'),
            (['in'], ULONG, 'MethodLength'),
            (['in'], LPVOID, 'MethodData'),
            (['in'], ULONG, 'DataLength'),
            (['in'], POINTER(ULONG), 'BytesReturned')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'KsEvent',
            (['in'], PKSEVENT, 'Event'),
            (['in'], ULONG, 'EventLength'),
            (['in'], LPVOID, 'EventData'),
            (['in'], ULONG, 'DataLength'),
            (['in'], POINTER(ULONG), 'BytesReturned')
        )
    )


KSMETHOD_TYPE_NONE = 0x00000000
KSMETHOD_TYPE_READ = 0x00000001
KSMETHOD_TYPE_WRITE = 0x00000002
KSMETHOD_TYPE_MODIFY = 0x00000003
KSMETHOD_TYPE_SOURCE = 0x00000004
KSMETHOD_TYPE_SEND = 0x00000001
KSMETHOD_TYPE_SETSUPPORT = 0x00000100
KSMETHOD_TYPE_BASICSUPPORT = 0x00000200
KSMETHOD_TYPE_TOPOLOGY = 0x10000000
