from .data_types import *
from comtypes.automation import IDispatch
from comtypes.typeinfo import IRecordInfo, tagTYPEDESC, tagARRAYDESC, tagSAFEARRAYBOUND
from comtypes.automation import _midlSAFEARRAY, VARIANT  # NOQA
import ctypes
import comtypes

from .propidl import DECIMAL, PROPERTYKEY, PROPVARIANT
from .mmreg import WAVEFORMATEXTENSIBLE, WAVEFORMATEX















class tagRemSNB(ctypes.Structure):
    _pack_ = 4
    _fields_ = [
        ('ulCntStr', ULONG),
        ('ulCntChar', ULONG),
        ('rgString', POINTER(USHORT)),
    ]


wireSNB = POINTER(tagRemSNB)

class _FLAGGED_WORD_BLOB(ctypes.Structure):
    _pack_ = 4
    _fields_ = [
        ('fFlags', ULONG),
        ('clSize', ULONG),
        ('asData', POINTER(USHORT)),
    ]


class _wireVARIANT(ctypes.Structure):
    pass


class _wireBRECORD(ctypes.Structure):
    _fields_ = [
        ('fFlags', ULONG),
        ('clSize', ULONG),
        ('pRecInfo', POINTER(IRecordInfo)),
        ('pRecord', POINTER(UBYTE)),
    ]


class _MIDL_IOleAutomationTypes_0001(ctypes.Union):
    class _wireSAFEARR_BSTR(ctypes.Structure):
        _fields_ = [
            ('Size', ULONG),
            ('aBstr', POINTER(POINTER(_FLAGGED_WORD_BLOB))),
        ]

    class _wireSAFEARR_UNKNOWN(ctypes.Structure):
        _fields_ = [
            ('Size', ULONG),
            ('apUnknown', POINTER(POINTER(comtypes.IUnknown))),
        ]

    class _wireSAFEARR_DISPATCH(ctypes.Structure):
        _fields_ = [
            ('Size', ULONG),
            ('apDispatch', POINTER(POINTER(IDispatch))),
        ]

    class _wireSAFEARR_VARIANT(ctypes.Structure):
        _fields_ = [
            ('Size', ULONG),
            ('aVariant', POINTER(POINTER(_wireVARIANT))),
        ]

    class _wireSAFEARR_BRECORD(ctypes.Structure):
        _fields_ = [
            ('Size', ULONG),
            ('aRecord', POINTER(POINTER(_wireBRECORD))),
        ]

    class _wireSAFEARR_HAVEIID(ctypes.Structure):
        _fields_ = [
            ('Size', ULONG),
            ('apUnknown', POINTER(POINTER(comtypes.IUnknown))),
            ('iid', GUID),
        ]

    class _BYTE_SIZEDARR(ctypes.Structure):
        _fields_ = [
            ('clSize', ULONG),
            ('pData', POINTER(UBYTE)),
        ]

    class _SHORT_SIZEDARR(ctypes.Structure):
        _fields_ = [
            ('clSize', ULONG),
            ('pData', POINTER(USHORT)),
        ]

    class _LONG_SIZEDARR(ctypes.Structure):
        _fields_ = [
            ('clSize', ULONG),
            ('pData', POINTER(ULONG)),
        ]

    class _HYPER_SIZEDARR(ctypes.Structure):
        _fields_ = [
            ('clSize', ULONG),
            ('pData', POINTER(LONGLONG)),
        ]

    _fields_ = [
        ('BstrStr', _wireSAFEARR_BSTR),
        ('UnknownStr', _wireSAFEARR_UNKNOWN),
        ('DispatchStr', _wireSAFEARR_DISPATCH),
        ('VariantStr', _wireSAFEARR_VARIANT),
        ('RecordStr', _wireSAFEARR_BRECORD),
        ('HaveIidStr', _wireSAFEARR_HAVEIID),
        ('ByteStr', _BYTE_SIZEDARR),
        ('WordStr', _SHORT_SIZEDARR),
        ('LongStr', _LONG_SIZEDARR),
        ('HyperStr', _HYPER_SIZEDARR),
    ]


class _wireSAFEARRAY(ctypes.Structure):
    pass


class __MIDL_IOleAutomationTypes_0004(ctypes.Union):
    _fields_ = [
        ('llVal', LONGLONG),
        ('lVal', INT),
        ('bVal', UBYTE),
        ('iVal', SHORT),
        ('fltVal', FLOAT),
        ('dblVal', DOUBLE),
        ('boolVal', VARIANT_BOOL),
        ('scode', SCODE),
        ('cyVal', LONGLONG),
        ('date', DOUBLE),
        ('bstrVal', POINTER(_FLAGGED_WORD_BLOB)),
        ('punkVal', POINTER(comtypes.IUnknown)),
        ('pdispVal', POINTER(IDispatch)),
        ('parray', POINTER(POINTER(_wireSAFEARRAY))),
        ('brecVal', POINTER(_wireBRECORD)),
        ('pbVal', POINTER(UBYTE)),
        ('piVal', POINTER(SHORT)),
        ('plVal', POINTER(INT)),
        ('pllVal', POINTER(LONGLONG)),
        ('pfltVal', POINTER(FLOAT)),
        ('pdblVal', POINTER(DOUBLE)),
        ('pboolVal', POINTER(VARIANT_BOOL)),
        ('pscode', POINTER(SCODE)),
        ('pcyVal', POINTER(LONGLONG)),
        ('pdate', POINTER(DOUBLE)),
        ('pbstrVal', POINTER(POINTER(_FLAGGED_WORD_BLOB))),
        ('ppunkVal', POINTER(POINTER(comtypes.IUnknown))),
        ('ppdispVal', POINTER(POINTER(IDispatch))),
        ('pparray', POINTER(POINTER(POINTER(_wireSAFEARRAY)))),
        ('pvarVal', POINTER(POINTER(_wireVARIANT))),
        ('cVal', UBYTE),
        ('uiVal', USHORT),
        ('ulVal', ULONG),
        ('ullVal', ULONGLONG),
        ('intVal', INT),
        ('uintVal', UINT),
        ('decVal', DECIMAL),
        ('pdecVal', POINTER(DECIMAL)),
        ('pcVal', POINTER(UBYTE)),
        ('puiVal', POINTER(USHORT)),
        ('pulVal', POINTER(ULONG)),
        ('pullVal', POINTER(ULONGLONG)),
        ('pintVal', POINTER(INT)),
        ('puintVal', POINTER(UINT)),
    ]


_wireVARIANT._fields_ = [
    ('clSize', ULONG),
    ('rpcReserved', ULONG),
    ('vt', USHORT),
    ('wReserved1', USHORT),
    ('wReserved2', USHORT),
    ('wReserved3', USHORT),
    ('DUMMYUNIONNAME', __MIDL_IOleAutomationTypes_0004),
]

_wireVARIANT._anonymous_ = ('DUMMYUNIONNAME',)


class __MIDL_IOleAutomationTypes_0006(ctypes.Union):
    _fields_ = [
        ('oInst', ULONG),
        ('lpvarValue', POINTER(VARIANT)),
    ]


class _wireSAFEARRAY_UNION(ctypes.Structure):
    _fields_ = [
        ('sfType', ULONG),
        ('u', _MIDL_IOleAutomationTypes_0001),
    ]

    _anonymous_ = ('u',)


_wireSAFEARRAY._fields_ = [
    ('cDims', USHORT),
    ('fFeatures', USHORT),
    ('cbElements', ULONG),
    ('cLocks', ULONG),
    ('uArrayStructs', _wireSAFEARRAY_UNION),
    ('rgsabound', POINTER(tagSAFEARRAYBOUND)),
]


wirePSAFEARRAY = POINTER(POINTER(_wireSAFEARRAY))






