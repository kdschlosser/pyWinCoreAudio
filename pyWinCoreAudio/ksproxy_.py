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
from .ks import PKSMETHOD
from .mmdeviceapi import IActivateAudioInterfaceAsyncOperation
from .enum import (
    AUDIO_STREAM_CATEGORY,
    AUDCLNT_STREAMOPTIONS
)
from .constant import (
    KSCATEGORY_QUALITY,
    STATIC_KSCATEGORY_QUALITY
)

from .ks import (
    KS_COMPRESSION,
    KS_FRAMING_RANGE,
    KS_FRAMING_RANGE_WEIGHTED,
    KSCORRELATED_TIME,
    KSRESOLUTION,
    KSSTATE,
    PKSSTREAMALLOCATOR_STATUS,
    PKSMULTIPLE_ITEM,
    KSPIN_INTERFACE,
    KSPIN_MEDIUM,
    KSPIN_COMMUNICATION,
    PKSALLOCATOR_FRAMING_EX,
    PKSPROPERTY,
    PKSEVENT
)


STATIC_IID_IMediaFilter = (0x56A86899, 0xAD4, 0x11CE, 0xB0, 0x3A, 0x0, 0x20, 0xAF, 0xB, 0xA7, 0x70)
STATIC_IID_IBaseFilter = (0x56A86895, 0xAD4, 0x11CE, 0xB0, 0x3A, 0x0, 0x20, 0xAF, 0xB, 0xA7, 0x70)
STATIC_IID_IPin = (0x56a86891, 0x0ad4, 0x11ce, 0xb0, 0x3a, 0x00, 0x20, 0xaf, 0x0b, 0xa7, 0x70)
STATIC_IID_IKsObject = (0x423c13a2, 0x2070, 0x11d0, 0x9e, 0xf7, 0x00, 0xaa, 0x00, 0xa2, 0x16, 0xa1)
STATIC_IID_IKsPinEx = (0x7bb38260, 0xd19c, 0x11d2, 0xb3, 0x8a, 0x00, 0xa0, 0xc9, 0x5e, 0xc2, 0x2e)
STATIC_IID_IKsPin = (0xb61178d1, 0xa2d9, 0x11cf, 0x9e, 0x53, 0x00, 0xaa, 0x00, 0xa2, 0x16, 0xa1)
STATIC_IID_IKsPinPipe = (0xe539cd90, 0xa8b4, 0x11d1, 0x81, 0x89, 0x00, 0xa0, 0xc9, 0x06, 0x28, 0x02)
STATIC_IID_IKsDataTypeHandler = (0x5ffbaa02, 0x49a3, 0x11d0, 0x9f, 0x36, 0x00, 0xaa, 0x00, 0xa2, 0x16, 0xa1)
STATIC_IID_IKsDataTypeCompletion = (0x827D1A0E, 0x0F73, 0x11D2, 0xB2, 0x7A, 0x00, 0xA0, 0xC9, 0x22, 0x31, 0x96)
STATIC_IID_IKsInterfaceHandler = (0xD3ABC7E0, 0x9A61, 0x11D0, 0xA4, 0x0D, 0x00, 0xA0, 0xC9, 0x22, 0x31, 0x96)
STATIC_IID_IKsClockPropertySet = (0x5C5CBD84, 0xE755, 0x11D0, 0xAC, 0x18, 0x00, 0xA0, 0xC9, 0x22, 0x31, 0x96)
STATIC_IID_IKsAllocator = (0x8da64899, 0xc0d9, 0x11d0, 0x84, 0x13, 0x00, 0x00, 0xf8, 0x22, 0xfe, 0x8a)
STATIC_IID_IKsAllocatorEx = (0x091bb63a, 0x603f, 0x11d1, 0xb0, 0x67, 0x00, 0xa0, 0xc9, 0x06, 0x28, 0x02)
STATIC_IID_IKsPropertySet = (0x31EFAC30, 0x515C, 0x11d0, 0xA9, 0xAA, 0x00, 0xAA, 0x00, 0x61, 0xBE, 0x93)
STATIC_IID_IKsTopology = (0x28F54683, 0x06FD, 0x11D2, 0xB2, 0x7A, 0x00, 0xA0, 0xC9, 0x22, 0x31, 0x96)
STATIC_IID_IKsControl = (0x28F54685, 0x06FD, 0x11D2, 0xB2, 0x7A, 0x00, 0xA0, 0xC9, 0x22, 0x31, 0x96)
STATIC_IID_IKsAggregateControl = (0x7F40EAC0, 0x3947, 0x11D2, 0x87, 0x4E, 0x00, 0xA0, 0xC9, 0x22, 0x31, 0x96)
STATIC_CLSID_Proxy = (0x17CCA71B, 0xECD7, 0x11D0, 0xB9, 0x08, 0x00, 0xA0, 0xC9, 0x22, 0x31, 0x96)


IID_IMediaFilter = DEFINE_GUIDEX(*STATIC_IID_IMediaFilter)
IID_IBaseFilter = DEFINE_GUIDEX(*STATIC_IID_IBaseFilter)
IID_IPin = DEFINE_GUIDEX(*STATIC_IID_IPin)
IID_IKsObject = DEFINE_GUIDEX(*STATIC_IID_IKsObject)
IID_IKsPin = DEFINE_GUIDEX(*STATIC_IID_IKsPinEx)
IID_IKsPinEx = DEFINE_GUIDEX(*STATIC_IID_IKsPin)
IID_IKsPinPipe = DEFINE_GUIDEX(*STATIC_IID_IKsPinPipe)
IID_IKsDataTypeHandler = DEFINE_GUIDEX(*STATIC_IID_IKsDataTypeHandler)
IID_IKsDataTypeCompletion = DEFINE_GUIDEX(*STATIC_IID_IKsDataTypeCompletion)
IID_IKsInterfaceHandler = DEFINE_GUIDEX(*STATIC_IID_IKsInterfaceHandler)
IID_IKsClockPropertySet = DEFINE_GUIDEX(*STATIC_IID_IKsClockPropertySet)
IID_IKsAllocator = DEFINE_GUIDEX(*STATIC_IID_IKsAllocator)
IID_IKsAllocatorEx = DEFINE_GUIDEX(*STATIC_IID_IKsAllocatorEx)

IID_IKsPropertySet = DEFINE_GUIDEX(*STATIC_IID_IKsPropertySet)
IID_IKsTopology = DEFINE_GUIDEX(*STATIC_IID_IKsTopology)
IID_IKsControl = DEFINE_GUIDEX(*STATIC_IID_IKsControl)
IID_IKsAggregateControl = DEFINE_GUIDEX(*STATIC_IID_IKsAggregateControl)
CLSID_Proxy = DEFINE_GUIDEX(*STATIC_CLSID_Proxy)


IID_IKsQualityForwarder = KSCATEGORY_QUALITY
STATIC_IID_IKsQualityForwarder = STATIC_KSCATEGORY_QUALITY


class KSALLOCATORMODE(ENUM):
    KsAllocatorMode_User = 0
    KsAllocatorMode_Kernel = 1


class FRAMING_PROP(ENUM):
    Framing_Prop_Uninitialized = 0
    Framing_Prop_None = 1
    Framing_Prop_Old = 2
    Framing_Prop_Ex = 3


PFRAMING_PROP = POINTER(FRAMING_PROP)


class FRAMING_CACHE_OPS(ENUM):
    Framing_Cache_Update = 0  # request to bypass cache when read/write
    Framing_Cache_ReadLast = 1
    Framing_Cache_ReadOrig = 2
    Framing_Cache_Write = 3


class OPTIMAL_WEIGHT_TOTALS(ctypes.Structure):
    _fields_ = [
        ('MinTotalNominator', LONGLONG),
        ('MaxTotalNominator', LONGLONG),
        ('TotalDenominator', LONGLONG)
    ]


# allocators strategy is defined by graph manager
AllocatorStrategy_DontCare = 0

# what to optimize
AllocatorStrategy_MinimizeNumberOfFrames = 0x00000001
AllocatorStrategy_MinimizeFrameSize = 0x00000002
AllocatorStrategy_MinimizeNumberOfAllocators = 0x00000004
AllocatorStrategy_MaximizeSpeed = 0x00000008


# factors (flags) defining the Pipes properties
PipeFactor_None = 0
PipeFactor_UserModeUpstream = 0x00000001
PipeFactor_UserModeDownstream = 0x00000002
PipeFactor_MemoryTypes = 0x00000004
PipeFactor_Flags = 0x00000008
PipeFactor_PhysicalRanges = 0x00000010
PipeFactor_OptimalRanges = 0x00000020
PipeFactor_FixedCompression = 0x00000040
PipeFactor_UnknownCompression = 0x00000080

PipeFactor_Buffers = 0x00000100
PipeFactor_Align = 0x00000200

PipeFactor_PhysicalEnd = 0x00000400
PipeFactor_LogicalEnd = 0x00000800


class IKsAllocator(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IKsClockPropertySet


class IKsAllocatorEx(IKsAllocator):
    _case_insensitive_ = False
    _iid_ = IID_IKsClockPropertySet


class IKsPin(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IKsClockPropertySet


class _KSSTREAM_SEGMENT(ctypes.Structure):
    pass


KSSTREAM_SEGMENT = _KSSTREAM_SEGMENT
PKSSTREAM_SEGMENT = POINTER(_KSSTREAM_SEGMENT)


class PIPE_STATE(ENUM):
    Pipe_State_DontCare = 0
    Pipe_State_RangeNotFixed = 1
    Pipe_State_RangeFixed = 2
    Pipe_State_CompressionUnknown = 3
    Pipe_State_Finalized = 4


# pipe dimensions relative to BeginPin.
class _PIPE_DIMENSIONS(ctypes.Structure):
    _fields_ = [
        ('AllocatorPin', KS_COMPRESSION),
        ('MaxExpansionPin', KS_COMPRESSION),
        ('EndPin', KS_COMPRESSION)
    ]


PIPE_DIMENSIONS = _PIPE_DIMENSIONS
PPIPE_DIMENSIONS = POINTER(_PIPE_DIMENSIONS)


class PIPE_ALLOCATOR_PLACE(ENUM):
    Pipe_Allocator_None = 0
    Pipe_Allocator_FirstPin = 1
    Pipe_Allocator_LastPin = 2
    Pipe_Allocator_MiddlePin = 3


PPIPE_ALLOCATOR_PLACE = POINTER(PIPE_ALLOCATOR_PLACE)


class KS_LogicalMemoryType(ENUM):
    KS_MemoryTypeDontCare = 0
    KS_MemoryTypeKernelPaged = 1
    KS_MemoryTypeKernelNonPaged = 2
    KS_MemoryTypeDeviceHostMapped = 3
    KS_MemoryTypeDeviceSpecific = 4
    KS_MemoryTypeUser = 5
    KS_MemoryTypeAnyHost = 6


PKS_LogicalMemoryType = POINTER(KS_LogicalMemoryType)


class _PIPE_TERMINATION(ctypes.Structure):
    _fields_ = [
        ('Flags', ULONG),
        ('OutsideFactors', ULONG),
        ('Weight', ULONG),  # outside weight
        ('PhysicalRange', KS_FRAMING_RANGE),
        ('OptimalRange', KS_FRAMING_RANGE_WEIGHTED),
        ('Compression', KS_COMPRESSION)  # relative to the connected pin on a neighboring filter.
    ]


PIPE_TERMINATION = _PIPE_TERMINATION


# extended allocator properties
class _ALLOCATOR_PROPERTIES_EX(ctypes.Structure):
    _fields_ = [
        ('cBuffers', LONG),
        ('cbBuffer', LONG),
        ('cbAlign', LONG),
        ('cbPrefix', LONG),
        # new part
        ('MemoryType', GUID),
        ('BusType', GUID),  # one of the buses this pipe is using
        ('State', PIPE_STATE),
        ('Input', PIPE_TERMINATION),
        ('Output', PIPE_TERMINATION),
        ('Strategy', ULONG),
        ('Flags', ULONG),
        ('Weight', ULONG),
        ('LogicalMemoryType', KS_LogicalMemoryType),
        ('AllocatorPlace', PIPE_ALLOCATOR_PLACE),
        ('Dimensions', PIPE_DIMENSIONS),
        ('PhysicalRange', KS_FRAMING_RANGE),  # on allocator pin
        ('PrevSegment', POINTER(IKsAllocatorEx)),  # doubly-linked list of KS allocators
        ('CountNextSegments', ULONG),  # possible multiple dependent pipes
        ('NextSegments', POINTER(POINTER(IKsAllocatorEx))),
        ('InsideFactors', ULONG),  # existing factors (different from "don't care")
        ('NumberPins', ULONG)
    ]


ALLOCATOR_PROPERTIES_EX = _ALLOCATOR_PROPERTIES_EX

PALLOCATOR_PROPERTIES_EX = POINTER(ALLOCATOR_PROPERTIES_EX)


class IKsClockPropertySet(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IKsClockPropertySet
    _methods_ = (
        comtypes.STDMETHOD(
            HRESULT,
            'KsGetTime',
            (POINTER(LONGLONG),)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsSetTime',
            (LONGLONG,)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsGetPhysicalTime',
            (POINTER(LONGLONG),)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsSetPhysicalTime',
            (LONGLONG,)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsGetCorrelatedTime',
            (POINTER(KSCORRELATED_TIME),)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsSetCorrelatedTime',
            (POINTER(KSCORRELATED_TIME),)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsGetCorrelatedPhysicalTime',
            (POINTER(KSCORRELATED_TIME),)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsSetCorrelatedPhysicalTime',
            (POINTER(KSCORRELATED_TIME),)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsGetResolution',
            (POINTER(KSRESOLUTION),)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsGetState',
            (POINTER(KSSTATE),)
        )
    )


IKsAllocator._methods_ = (
        comtypes.STDMETHOD(
            HANDLE,
            'KsGetAllocatorHandle',
            ()
        ),
        comtypes.STDMETHOD(
            KSALLOCATORMODE,
            'KsGetAllocatorMode',
            ()
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsGetAllocatorStatus',
            (PKSSTREAMALLOCATOR_STATUS,)
        ),
        comtypes.STDMETHOD(
            VOID,
            'KsSetAllocatorMode',
            (KSALLOCATORMODE,)
        )
    )


IKsAllocatorEx._methods_ = (
        comtypes.STDMETHOD(
            PALLOCATOR_PROPERTIES_EX,
            'KsGetProperties',
            ()
        ),
        comtypes.STDMETHOD(
            VOID,
            'KsSetProperties',
            (PALLOCATOR_PROPERTIES_EX,)
        ),
        comtypes.STDMETHOD(
            VOID,
            'KsSetAllocatorHandle',
            (HANDLE,)
        ),
        comtypes.STDMETHOD(
            HANDLE,
            'KsCreateAllocatorAndGetHandle',
            (POINTER(IKsPin),)
        )
    )


class KSPEEKOPERATION(ENUM):
    KsPeekOperation_PeekOnly = 0
    KsPeekOperation_AddRef = 1


IKsPin._methods_ = (
        comtypes.STDMETHOD(
            HRESULT,
            'KsQueryMediums',
            (POINTER(PKSMULTIPLE_ITEM),)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsQueryInterfaces',
            (POINTER(PKSMULTIPLE_ITEM),)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsCreateSinkPinHandle',
            (KSPIN_INTERFACE, KSPIN_MEDIUM)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsGetCurrentCommunication',
            (POINTER(KSPIN_COMMUNICATION), POINTER(KSPIN_INTERFACE), POINTER(KSPIN_MEDIUM))
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsPropagateAcquire',
            ()
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsDeliver',
            (POINTER(VOID), ULONG)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsMediaSamplesCompleted',
            (PKSSTREAM_SEGMENT,)
        ),
        comtypes.STDMETHOD(
            POINTER(VOID),
            'KsPeekAllocator',
            (KSPEEKOPERATION,)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsReceiveAllocator',
            (POINTER(VOID),)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsRenegotiateAllocator',
            ()
        ),
        comtypes.STDMETHOD(
            LONG,
            'KsIncrementPendingIoCount',
            ()
        ),
        comtypes.STDMETHOD(
            LONG,
            'KsDecrementPendingIoCount',
            ()
        ),
        comtypes.STDMETHOD(
            VOID,
            'KsQualityNotify',
            (ULONG, REFERENCE_TIME)
        )
    )


class IKsPinEx(IKsPin):
    _case_insensitive_ = False
    _iid_ = IID_IKsClockPropertySet
    _methods_ = (
        comtypes.STDMETHOD(
            VOID,
            'KsNotifyError',
            (POINTER(VOID), HRESULT)
        ),
    )


class _PIN_DIRECTION(ENUM):
    PINDIR_INPUT = 0
    PINDIR_OUTPUT = PINDIR_INPUT + 1


PIN_DIRECTION = _PIN_DIRECTION


class IMediaFilter(comtypes.IPersist):
    _case_insensitive_ = False
    _iid_ = IID_IMediaFilter


IMediaFilter._methods_ = (
    comtypes.STDMETHOD(
        HRESULT,
        'Stop',
        ()
    ),
    comtypes.STDMETHOD(
        HRESULT,
        'Pause',
        ()
    ),
    comtypes.STDMETHOD(
        HRESULT,
        'Run',
        (REFERENCE_TIME,)
    ),
    comtypes.STDMETHOD(
        HRESULT,
        'GetState',
        (DWORD, POINTER(VOID),)
    ),
    comtypes.STDMETHOD(
        HRESULT,
        'SetSyncSource',
        (POINTER(VOID),)
    ),
    comtypes.STDMETHOD(
        HRESULT,
        'GetSyncSource',
        (POINTER(POINTER(VOID)),)
    )
)


class IBaseFilter(IMediaFilter):
    _case_insensitive_ = False
    _iid_ = IID_IBaseFilter


class _PIN_INFO(ctypes.Structure):
    _fields_ = [
        ('pFilter', POINTER(IBaseFilter)),
        ('dir', PIN_DIRECTION),
        ('achName', WCHAR * 128)
    ]


PIN_INFO = _PIN_INFO


class _AM_MEDIA_TYPE(ctypes.Structure):
    _fields_ = [
        ('majortype', GUID),
        ('subtype', GUID),
        ('bFixedSizeSamples', WINBOOL),
        ('bTemporalCompression', WINBOOL),
        ('lSampleSize', ULONG),
        ('formattype', GUID),
        ('pUnk', POINTER(comtypes.IUnknown)),
        ('cbFormat', ULONG),
        ('pbFormat', POINTER(BYTE))
    ]


AM_MEDIA_TYPE = _AM_MEDIA_TYPE


class IPin(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IPin


IBaseFilter._methods_ = (
    comtypes.STDMETHOD(
        HRESULT,
        'EnumPins',
        (POINTER(POINTER(VOID)),)
    ),
    comtypes.STDMETHOD(
        HRESULT,
        'FindPin',
        (LPCWSTR, POINTER(POINTER(IPin)))
    ),
    comtypes.STDMETHOD(
        HRESULT,
        'QueryFilterInfo',
        (POINTER(VOID),)
    ),
    comtypes.STDMETHOD(
        HRESULT,
        'JoinFilterGraph',
        (POINTER(VOID), LPCWSTR)
    ),
    comtypes.STDMETHOD(
        HRESULT,
        'QueryVendorInfo',
        (POINTER(LPWSTR), )
    )
)
    

IPin._methods_ = (
        comtypes.STDMETHOD(
            HRESULT,
            'Connect',
            (POINTER(IPin), POINTER(AM_MEDIA_TYPE))
        ),    
        comtypes.STDMETHOD(
            HRESULT,
            'ReceiveConnection',
            (POINTER(IPin), POINTER(AM_MEDIA_TYPE))
        ),    
        comtypes.STDMETHOD(
            HRESULT,
            'Disconnect',
            ()
        ),    
        comtypes.STDMETHOD(
            HRESULT,
            'ConnectedTo',
            (POINTER(POINTER(IPin)),)
        ),    
        comtypes.STDMETHOD(
            HRESULT,
            'ConnectionMediaType',
            (POINTER(AM_MEDIA_TYPE),)
        ),    
        comtypes.STDMETHOD(
            HRESULT,
            'QueryPinInfo',
            (POINTER(PIN_INFO),)
        ),    
        comtypes.STDMETHOD(
            HRESULT,
            'QueryDirection',
            (POINTER(PIN_DIRECTION),)
        ),    
        comtypes.STDMETHOD(
            HRESULT,
            'QueryId',
            (POINTER(LPWSTR),)
        ),    
        comtypes.STDMETHOD(
            HRESULT,
            'QueryAccept',
            (POINTER(AM_MEDIA_TYPE),)
        ),    
        comtypes.STDMETHOD(
            HRESULT,
            'EnumMediaTypes',
            (POINTER(POINTER(VOID)),)
        ),    
        comtypes.STDMETHOD(
            HRESULT,
            'QueryInternalConnections',
            (POINTER(POINTER(IPin)), POINTER(ULONG), FRAMING_CACHE_OPS)
        ),    
        comtypes.STDMETHOD(
            HRESULT,
            'EndOfStream',
            ()
        ),    
        comtypes.STDMETHOD(
            HRESULT,
            'BeginFlush',
            ()
        ),    
        comtypes.STDMETHOD(
            HRESULT,
            'EndFlush',
            ()
        ),    
        comtypes.STDMETHOD(
            HRESULT,
            'NewSegment',
            (REFERENCE_TIME, REFERENCE_TIME, DOUBLE)
        )
    )


class IKsPinPipe(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IKsClockPropertySet
    _methods_ = (
        comtypes.STDMETHOD(
            HRESULT,
            'KsGetPinFramingCache',
            (POINTER(PKSALLOCATOR_FRAMING_EX), PFRAMING_PROP, FRAMING_CACHE_OPS)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsSetPinFramingCache',
            (PKSALLOCATOR_FRAMING_EX, PFRAMING_PROP, FRAMING_CACHE_OPS)
        ),
        comtypes.STDMETHOD(
            POINTER(IPin),
            'KsGetConnectedPin',
            ()
        ),
        comtypes.STDMETHOD(
            POINTER(IKsAllocatorEx),
            'KsGetPipe',
            (KSPEEKOPERATION,)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsSetPipe',
            (POINTER(IKsAllocatorEx),)
        ),
        comtypes.STDMETHOD(
            ULONG,
            'KsGetPipeAllocatorFlag',
            ()
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsSetPipeAllocatorFlag',
            (ULONG,)
        ),
        comtypes.STDMETHOD(
            GUID,
            'KsGetPinBusCache',
            ()
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsSetPinBusCache',
            (GUID,)
        ),
        # very useful methods for tracing.
        comtypes.STDMETHOD(
            PWCHAR,
            'KsGetPinName',
            ()
        ),
        comtypes.STDMETHOD(
            PWCHAR,
            'KsGetFilterName',
            ()
        )
    )


class IKsPinFactory(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IKsClockPropertySet
    _methods_ = (
        comtypes.STDMETHOD(
            HRESULT,
            'KsPinFactory',
            (POINTER(ULONG),)
        )
    )


class KSIOOPERATION(ENUM):
    KsIoOperation_Write = 0
    KsIoOperation_Read = 1


class IKsDataTypeHandler(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IKsClockPropertySet
    _methods_ = (
        comtypes.STDMETHOD(
            HRESULT,
            'KsCompleteIoOperation',
            (POINTER(VOID), PVOID, KSIOOPERATION, BOOL)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsIsMediaTypeInRanges',
            (PVOID,)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsPrepareIoOperation',
            (POINTER(VOID), PVOID, KSIOOPERATION)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsQueryExtendedSize',
            (POINTER(ULONG),)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsSetMediaType',
            (POINTER(AM_MEDIA_TYPE),)
        )
    )


class IKsDataTypeCompletion(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IKsClockPropertySet
    _methods_ = (
        comtypes.STDMETHOD(
            HRESULT,
            'KsCompleteMediaType',
            (HANDLE, ULONG, POINTER(AM_MEDIA_TYPE))
        )
    )


class IKsInterfaceHandler(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IKsClockPropertySet
    _methods_ = (
        comtypes.STDMETHOD(
            HRESULT,
            'KsSetPin',
            (POINTER(IKsPin),)
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsProcessMediaSamples',
            (
                POINTER(IKsDataTypeHandler),
                POINTER(POINTER(VOID)),
                PLONG,
                KSIOOPERATION,
                POINTER(PKSSTREAM_SEGMENT)
            )
        ),
        comtypes.STDMETHOD(
            HRESULT,
            'KsCompleteIo',
            (
                PKSSTREAM_SEGMENT,
            )
        )
    )


# This structure definition is the common header required by the proxy to
# dispatch the stream segment to the interface handler.  Interface handlers
# will create extended structures to include other information such as
# media samples, extended header size and so on.
_KSSTREAM_SEGMENT._fields_ = [
        ('KsInterfaceHandler', POINTER(IKsInterfaceHandler)),
        ('KsDataTypeHandler', POINTER(IKsDataTypeHandler)),
        ('IoOperation', KSIOOPERATION),
        ('CompletionEvent', HANDLE)
    ]


class IKsObject(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IKsClockPropertySet
    _methods_ = (
        comtypes.STDMETHOD(
            HANDLE,
            'KsGetObjectHandle',
            ()
        )
    )


class IKsQualityForwarder(IKsObject):
    _case_insensitive_ = False
    _iid_ = IID_IKsClockPropertySet
    _methods_ = (
        comtypes.STDMETHOD(
            VOID,
            'KsFlushClient',
            (POINTER(IKsPin),)
        )
    )


class IKsNotifyEvent(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IKsClockPropertySet
    _methods_ = (
        comtypes.STDMETHOD(
            HRESULT,
            'KsNotifyEvent',
            (ULONG, ULONG_PTR, ULONG_PTR)
        )
    )

#
# KSDDKAPI HRESULT WINAPI
# KsResolveRequiredAttributes(
#     PKSDATARANGE DataRange,
#     PKSMULTIPLE_ITEM Attributes OPTIONAL
#     );
#
# KSDDKAPI HRESULT WINAPI
# KsOpenDefaultDevice(
#     REFGUID Category,
#     ACCESS_MASK Access,
#     PHANDLE DeviceHandle
#     );
#
# KSDDKAPI HRESULT WINAPI
# KsSynchronousDeviceControl(
#     HANDLE      Handle,
#     ULONG       IoControl,
#     PVOID       InBuffer,
#     ULONG       InLength,
#     PVOID       OutBuffer,
#     ULONG       OutLength,
#     PULONG      BytesReturned
#     );
#
# KSDDKAPI HRESULT WINAPI
# KsGetMultiplePinFactoryItems(
#     HANDLE  FilterHandle,
#     ULONG   PinFactoryId,
#     ULONG   PropertyId,
#     PVOID*  Items
#     );
#
#
# KSDDKAPI HRESULT WINAPI
# KsGetMediaTypeCount(
#     HANDLE      FilterHandle,
#     ULONG       PinFactoryId,
#     ULONG*      MediaTypeCount
#     );
#
#
# KSDDKAPI HRESULT WINAPI
# KsGetMediaType(
#     int         Position,
#     AM_MEDIA_TYPE* AmMediaType,
#     HANDLE      FilterHandle,
#     ULONG       PinFactoryId
#     );
#


KSPROPERTY_SUPPORT_GET = 1
KSPROPERTY_SUPPORT_SET = 2


class IKsPropertySet(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IKsClockPropertySet
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'Set',
            (['in'], REFGUID, 'PropSet'),
            (['in'], ULONG, 'Id'),
            (['in'], LPVOID, 'InstanceData'),
            (['in'], ULONG, 'InstanceLength'),
            (['in'], LPVOID, 'PropertyData'),
            (['in'], ULONG, 'DataLength')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'Get',
            (['in'], REFGUID, 'PropSet'),
            (['in'], ULONG, 'Id'),
            (['in'], LPVOID, 'InstanceData'),
            (['in'], ULONG, 'InstanceLength'),
            (['out'], LPVOID, 'PropertyData'),
            (['in'], ULONG, 'DataLength'),
            (['out'], POINTER(ULONG), 'BytesReturned')

        ),
        COMMETHOD(
            [],
            HRESULT,
            'QuerySupported',
            (['in'], REFGUID, 'PropSet'),
            (['in'], ULONG, 'Id'),
            (['out'], POINTER(ULONG), 'TypeSupport')
        )
    )


class IKsControl(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IKsClockPropertySet
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


class IKsAggregateControl(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IKsClockPropertySet
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'KsAddAggregate',
            (['in'], REFGUID, 'AggregateClass')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'KsRemoveAggregate',
            (['in'], REFGUID, 'AggregateClass')
        )
    )


class IKsTopology(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IKsClockPropertySet
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'CreateNodeInstance',
            (['in'], ULONG, 'NodeId'),
            (['in'], ULONG, 'Flags'),
            (['in'], ACCESS_MASK, 'DesiredAccess'),
            (['in'], POINTER(comtypes.IUnknown), 'UnkOuter'),
            (['in'], REFGUID, 'InterfaceId'),
            (['out'], POINTER(LPVOID), 'Interface')
        )
    )


PIUnknown = POINTER(comtypes.IUnknown)


class WAVEFORMATEX(ctypes.Structure):
    _fields_ = [
        ('wFormatTag', WORD),
        ('nChannels', WORD),
        ('nSamplesPerSec', WORD),
        ('nAvgBytesPerSec', WORD),
        ('nBlockAlign', WORD),
        ('wBitsPerSample', WORD),
        ('cbSize', WORD),
    ]


PWAVEFORMATEX = POINTER(WAVEFORMATEX)


class AudioClientProperties(ctypes.Structure):
    _fields_ = [
        ('cbSize', UINT32),
        ('bIsOffload', BOOL),
        ('eCategory', AUDIO_STREAM_CATEGORY),
        ('Options', AUDCLNT_STREAMOPTIONS),
    ]


PAudioClientProperties = POINTER(AudioClientProperties)


class IKsControl(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IKsControl
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'KsEvent',
            (['in'], ULONG, 'EventLength'),
            (['in', 'out'], LPVOID, 'EventData'),
            (['in'], ULONG, 'DataLength'),
            (['in', 'out'], LPULONG, 'BytesReturned')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'KsMethod',
            (['in'], PKSMETHOD, 'Method'),
            (['in'], ULONG, 'MethodLength'),
            (['in', 'out'], LPVOID, 'MethodData'),
            (['in'], ULONG, 'DataLength'),
            (['in', 'out'], LPULONG, 'BytesReturned')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'KsProperty',
            (['out'], LPHRESULT, 'activateOperation'),
            (['out'], POINTER(PIUnknown), 'activatedInterface')
        ),
    )


PIActivateAudioInterfaceAsyncOperation = POINTER(
    IActivateAudioInterfaceAsyncOperation
)
