from .data_types import *

import datetime
import decimal
import array
from ctypes import _Pointer  # NOQA
from _ctypes import CopyComPointer

from comtypes import _safearray  # NOQA
from comtypes._safearray import SAFEARRAY  # NOQA
from comtypes import npsupport
from comtypes.typeinfo import IRecordInfo


from comtypes.automation import (
    DECIMAL,
    CY,
    IDispatch,
    SCODE,
    VT_EMPTY,
    VT_NULL,
    VT_I2,
    VT_I4,
    VT_R4,
    VT_R8,
    VT_CY,
    VT_DATE,
    VT_BSTR,
    VT_DISPATCH,
    VT_ERROR,
    VT_BOOL,
    VT_VARIANT,
    VT_UNKNOWN,
    VT_DECIMAL,
    VT_I1,
    VT_UI1,
    VT_UI2,
    VT_UI4,
    VT_I8,
    VT_UI8,
    VT_INT,
    VT_UINT,
    VT_VOID,  # NOQA
    VT_HRESULT,  # NOQA
    VT_PTR,  # NOQA
    VT_SAFEARRAY,  # NOQA
    VT_CARRAY,  # NOQA
    VT_USERDEFINED,  # NOQA
    VT_LPSTR,
    VT_LPWSTR,
    VT_RECORD,  # NOQA
    VT_INT_PTR,  # NOQA
    VT_UINT_PTR,  # NOQA
    VT_FILETIME,
    VT_BLOB,
    VT_STREAM,
    VT_STORAGE,
    VT_STREAMED_OBJECT,
    VT_STORED_OBJECT,
    VT_BLOB_OBJECT,  # NOQA
    VT_CF,
    VT_CLSID,
    VT_VERSIONED_STREAM,
    VT_BSTR_BLOB,
    VT_VECTOR,
    VT_ARRAY,
    VT_BYREF,
    VT_RESERVED,  # NOQA
    VT_ILLEGAL,  # NOQA
    VT_ILLEGALMASKED,  # NOQA
    VT_TYPEMASK,  # NOQA
    _com_null_date,  # NOQA
    _SysAllocStringLen,  # NOQA
    _midlSAFEARRAY,  # NOQA
    _vartype_to_ctype,  # NOQA
    _ctype_to_vartype,  # NOQA
    _arraycode_to_vartype,  # NOQA
    _byref_type  # NOQA
)

import ctypes
from comtypes import (
    IUnknown,
    helpstring,
    COMMETHOD
)

LPSAFEARRAY = POINTER(SAFEARRAY)

PROPID = ULONG

VARIANT_TRUE = 0xFFFF
VARIANT_FALSE = 0x0000


class tagVARIANT(ctypes.Structure):
    pass


VARIANT = tagVARIANT


# noinspection PyAttributeOutsideInit
class tagPROPVARIANT(ctypes.Structure):

    @property
    def cVal(self):
        """
        c_char
        VT_I1
        """
        return self._cVal

    @cVal.setter
    def cVal(self, value):
        _PropVariantClear(self)
        self._cVal = value
        self.vt = VT_I1

    @cVal.deleter
    def cVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def bVal(self):
        """
        c_ubyte
        VT_UI1
        """
        return self._bVal

    @bVal.setter
    def bVal(self, value):
        _PropVariantClear(self)
        self._bVal = value
        self.vt = VT_UI1

    @bVal.deleter
    def bVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def iVal(self):
        """
        c_short
        VT_I2
        """
        return self._iVal

    @iVal.setter
    def iVal(self, value):
        _PropVariantClear(self)
        self._iVal = value
        self.vt = VT_I2

    @iVal.deleter
    def iVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def uiVal(self):
        """
        c_ushort
        VT_UI2
        """
        return self._uiVal

    @uiVal.setter
    def uiVal(self, value):
        _PropVariantClear(self)
        self._uiVal = value
        self.vt = VT_UI2

    @uiVal.deleter
    def uiVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def lVal(self):
        """
        c_long
        VT_I4
        """
        return self._lVal

    @lVal.setter
    def lVal(self, value):
        _PropVariantClear(self)
        self._lVal = value
        self.vt = VT_I4

    @lVal.deleter
    def lVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def ulVal(self):
        """
        c_ulong
        VT_UI4
        """
        return self._ulVal

    @ulVal.setter
    def ulVal(self, value):
        _PropVariantClear(self)
        self._ulVal = value
        self.vt = VT_UI4

    @ulVal.deleter
    def ulVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def intVal(self):
        """
        c_int
        VT_INT
        """
        return self._intVal

    @intVal.setter
    def intVal(self, value):
        _PropVariantClear(self)
        self._intVal = value
        self.vt = VT_INT

    @intVal.deleter
    def intVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def uintVal(self):
        """
        c_uint
        VT_UINT
        """
        return self._uintVal

    @uintVal.setter
    def uintVal(self, value):
        _PropVariantClear(self)
        self._uintVal = value
        self.vt = VT_UINT

    @uintVal.deleter
    def uintVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def hVal(self):
        """
        LARGE_INTEGER
        VT_I8
        """
        return self._hVal

    @hVal.setter
    def hVal(self, value):
        _PropVariantClear(self)
        self._hVal = value
        self.vt = VT_I8

    @hVal.deleter
    def hVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def uhVal(self):
        """
        ULARGE_INTEGER
        VT_UI8
        """
        return self._uhVal

    @uhVal.setter
    def uhVal(self, value):
        _PropVariantClear(self)
        self._uhVal = value
        self.vt = VT_UI8

    @uhVal.deleter
    def uhVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def fltVal(self):
        """
        c_float
        VT_R4
        """
        return self._fltVal

    @fltVal.setter
    def fltVal(self, value):
        _PropVariantClear(self)
        self._fltVal = value
        self.vt = VT_R4

    @fltVal.deleter
    def fltVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def dblVal(self):
        """
        c_double
        VT_R8
        """
        return self._dblVal

    @dblVal.setter
    def dblVal(self, value):
        _PropVariantClear(self)
        self._dblVal = value
        self.vt = VT_R8

    @dblVal.deleter
    def dblVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def boolVal(self):
        """
        VARIANT_BOOL
        VT_BOOL
        """
        return self._boolVal.value

    @boolVal.setter
    def boolVal(self, value):
        _PropVariantClear(self)

        boolVal = VARIANT_BOOL()

        if value:
            boolVal.value = True
        else:
            boolVal.value = False

        self._boolVal = boolVal

        self.vt = VT_BOOL

    @boolVal.deleter
    def boolVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def scode(self):
        """
        SCODE
        VT_ERROR
        """
        return self._scode

    @scode.setter
    def scode(self, value):
        _PropVariantClear(self)
        self._scode = value
        self.vt = VT_ERROR

    @scode.deleter
    def scode(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def cyVal(self):
        """
        CY
        VT_CY
        """
        return self._cyVal / decimal.Decimal("10000")

    @cyVal.setter
    def cyVal(self, value):
        _PropVariantClear(self)
        self._cyVal = value
        self.vt = VT_CY

    @cyVal.deleter
    def cyVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def date(self):
        """
        DATE
        VT_DATE
        """
        days = self._date
        return datetime.timedelta(days=days) + _com_null_date

    @date.setter
    def date(self, value):
        _PropVariantClear(self)

        if npsupport.isdatetime64(value):
            com_days = value - npsupport.com_null_date64
            com_days /= npsupport.numpy.timedelta64(1, 'D')
            self._date = com_days
        else:
            delta = value - _com_null_date
            # a day has 24 * 60 * 60 = 86400 seconds
            com_days = delta.days + (
                    delta.seconds + delta.microseconds * 1e-6) / 86400.
            self._date = com_days

        self.vt = VT_DATE

    @date.deleter
    def date(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def filetime(self):
        """
        FILETIME
        VT_FILETIME
        """
        return self._filetime

    @filetime.setter
    def filetime(self, value):
        _PropVariantClear(self)
        self._filetime = value
        self.vt = VT_FILETIME

    @filetime.deleter
    def filetime(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def puuid(self):
        """
        POINTER(CLSID)
        VT_CLSID
        """
        return self._puuid

    @puuid.setter
    def puuid(self, value):
        _PropVariantClear(self)
        self._puuid = value
        self.vt = VT_CLSID

    @puuid.deleter
    def puuid(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pclipdata(self):
        """
        POINTER(CLIPDATA)
        VT_CF
        """
        return self._pclipdata

    @pclipdata.setter
    def pclipdata(self, value):
        _PropVariantClear(self)
        self._pclipdata = value
        self.vt = VT_CF

    @pclipdata.deleter
    def pclipdata(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def bstrVal(self):
        """
        BSTR
        VT_BSTR
        """
        return self._bstrVal

    @bstrVal.setter
    def bstrVal(self, value):
        _PropVariantClear(self)
        self._bstrVal = _SysAllocStringLen(value, len(value))
        self.vt = VT_BSTR

    @bstrVal.deleter
    def bstrVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def bstrblobVal(self):
        """
        BSTRBLOB
        VT_BSTR_BLOB
        """
        return self._bstrblobVal

    @bstrblobVal.setter
    def bstrblobVal(self, value):
        _PropVariantClear(self)
        self._bstrblobVal = value
        self.vt = VT_BSTR_BLOB

    @bstrblobVal.deleter
    def bstrblobVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def blob(self):
        """
        BLOB
        VT_BLOB, VT_BLOBOBJECT
        """

        res = []

        blob = self._blob
        for i in range(blob.cbSize):
            res.append(blob.pData[i])

        return res

    @blob.setter
    def blob(self, value):
        _PropVariantClear(self)

        blob = BLOB()

        try:
            blob.cbSize = len(value)
            pData = (BYTE * len(value))()
            for i, item in enumerate(value):
                pData[i] = item

            blob.pData = pData

            self._blob = blob
            self.vt = VT_BLOB

        except:  # NOQA
            raise ValueError('value must be an iterable that impliments __len__ and is filled with bytes')

    @blob.deleter
    def blob(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pszVal(self):
        """
        LPSTR
        VT_LPSTR
        """
        return self._pszVal

    @pszVal.setter
    def pszVal(self, value):
        _PropVariantClear(self)
        self._pszVal = value
        self.vt = VT_LPSTR

    @pszVal.deleter
    def pszVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pwszVal(self):
        """
        LPWSTR
        VT_LPWSTR
        """
        return self._pwszVal

    @pwszVal.setter
    def pwszVal(self, value):
        _PropVariantClear(self)
        self._pwszVal = value
        self.vt = VT_LPWSTR

    @pwszVal.deleter
    def pwszVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def punkVal(self):
        """
        POINTER(IUnknown)
        VT_UNKNOWN
        """

        val = self._punkVal
        if not val:
            # We should/could return a NULL COM pointer.
            # But the code generation must be able to construct one
            # from the __repr__ of it.
            return None  # XXX?

        ptr = ctypes.cast(val, POINTER(IUnknown))
        # ctypes.cast doesn't call AddRef (it should, imo!)
        ptr.AddRef()

        return ptr.__ctypes_from_outparam__()

    @punkVal.setter
    def punkVal(self, value):
        _PropVariantClear(self)
        CopyComPointer(value, ctypes.byref(self._punkVal))
        self.vt = VT_UNKNOWN

    @punkVal.deleter
    def punkVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pdispVal(self):
        """
        POINTER(IDispatch)
        VT_DISPATCH
        """
        val = self._pdispVal
        if not val:
            # See above.
            return None  # XXX?

        ptr = ctypes.cast(val, POINTER(IDispatch))
        # ctypes.cast doesn't call AddRef (it should, imo!)
        ptr.AddRef()

        # if not dynamic:
        return ptr.__ctypes_from_outparam__()
        # else:
        #     from comtypes.client.dynamic import Dispatch
        #     return Dispatch(ptr)

    @pdispVal.setter
    def pdispVal(self, value):
        _PropVariantClear(self)
        CopyComPointer(value, ctypes.byref(self._pdispVal))
        self.vt = VT_DISPATCH

    @pdispVal.deleter
    def pdispVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pStream(self):
        """
        POINTER(IStream)
        VT_STREAM, VT_STREAMED_OBJECT
        """
        return self._pStream

    @pStream.setter
    def pStream(self, value):
        _PropVariantClear(self)
        self._pStream = value
        self.vt = VT_STREAM, VT_STREAMED_OBJECT

    @pStream.deleter
    def pStream(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pStorage(self):
        """
        POINTER(IStorage)
        VT_STORAGE, VT_STORED_OBJECT
        """
        return self._pStorage

    @pStorage.setter
    def pStorage(self, value):
        _PropVariantClear(self)
        self._pStorage = value
        self.vt = VT_STORAGE, VT_STORED_OBJECT

    @pStorage.deleter
    def pStorage(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pVersionedStream(self):
        """
        LPVERSIONEDSTREAM
        VT_VERSIONED_STREAM
        """
        return self._pVersionedStream

    @pVersionedStream.setter
    def pVersionedStream(self, value):
        _PropVariantClear(self)
        self._pVersionedStream = value
        self.vt = VT_VERSIONED_STREAM

    @pVersionedStream.deleter
    def pVersionedStream(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def parray(self):
        """
        POINTER(_safearray.SAFEARRAY)
        VT_ARRAY
        """

        typ = _vartype_to_ctype[self.vt & ~VT_ARRAY]
        return ctypes.cast(self._parray, _midlSAFEARRAY(typ)).unpack()

    @parray.setter
    def parray(self, value):
        _PropVariantClear(self)

        if isinstance(value, (list, tuple)):
            obj = _midlSAFEARRAY(PROPVARIANT).create(value)

        elif isinstance(value, array.array):
            vartype = _arraycode_to_vartype[value.typecode]
            typ = _vartype_to_ctype[vartype]
            obj = _midlSAFEARRAY(typ).create(value)

        elif npsupport.isndarray(value):
            # Try to convert a simple array of basic types.
            descr = value.dtype.descr[0][1]
            typ = npsupport.typecodes.get(descr)
            if typ is None:
                # Try for variant
                obj = _midlSAFEARRAY(PROPVARIANT).create(value)
            else:
                obj = _midlSAFEARRAY(typ).create(value)
        else:
            raise TypeError('Unsupported data type ({0})'.format(type(value)))

        ctypes.memmove(
            ctypes.byref(self._parray),
            ctypes.byref(obj),
            ctypes.sizeof(obj)
        )
        self.vt = VT_ARRAY | obj._vartype_  # NOQA

    @parray.deleter
    def parray(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def cac(self):
        """
        CAC
        VT_VECTOR  |  VT_I1
        """
        return self._cac

    @cac.setter
    def cac(self, value):
        _PropVariantClear(self)
        self._cac = value
        self.vt = VT_VECTOR | VT_I1

    @cac.deleter
    def cac(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def caub(self):
        """
        CAUB
        VT_VECTOR  |  VT_UI1
        """
        res = []

        caub = self._caub

        count = caub.cElems / ctypes.sizeof(UBYTE)
        for i in range(count):
            res.append(caub.pElems[i])

        return res

    @caub.setter
    def caub(self, value):
        _PropVariantClear(self)

        caub = CAUB()

        try:
            caub.cElems = len(value) * ctypes.sizeof(UBYTE)
            pElems = (UBYTE * len(value))()
            for i, item in enumerate(value):
                pElems[i] = item

            caub.pElems = pElems

            self._caub = caub
            self.vt = VT_VECTOR | VT_UI1

        except:  # NOQA
            raise ValueError(
                'value must be an iterable that impliments __len__ and is filled with unsigned bytes'
            )

    @caub.deleter
    def caub(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def cai(self):
        """
        CAI
        VT_VECTOR  |  VT_I2
        """
        return self._cai

    @cai.setter
    def cai(self, value):
        _PropVariantClear(self)
        self._cai = value
        self.vt = VT_VECTOR | VT_I2

    @cai.deleter
    def cai(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def caui(self):
        """
        CAUI
        VT_VECTOR  |  VT_UI2
        """
        return self._caui

    @caui.setter
    def caui(self, value):
        _PropVariantClear(self)
        self._caui = value
        self.vt = VT_VECTOR | VT_UI2

    @caui.deleter
    def caui(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def cal(self):
        """
        CAL
        VT_VECTOR  |  VT_I4
        """
        return self._cal

    @cal.setter
    def cal(self, value):
        _PropVariantClear(self)
        self._cal = value
        self.vt = VT_VECTOR | VT_I4

    @cal.deleter
    def cal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def caul(self):
        """
        CAUL
        VT_VECTOR  |  VT_UI4
        """
        return self._caul

    @caul.setter
    def caul(self, value):
        _PropVariantClear(self)
        self._caul = value
        self.vt = VT_VECTOR | VT_UI4

    @caul.deleter
    def caul(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def cah(self):
        """
        CAH
        VT_VECTOR  |  VT_I8
        """
        return self._cah

    @cah.setter
    def cah(self, value):
        _PropVariantClear(self)
        self._cah = value
        self.vt = VT_VECTOR | VT_I8

    @cah.deleter
    def cah(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def cauh(self):
        """
        CAUH
        VT_VECTOR  |  VT_UI8
        """
        return self._cauh

    @cauh.setter
    def cauh(self, value):
        _PropVariantClear(self)
        self._cauh = value
        self.vt = VT_VECTOR | VT_UI8

    @cauh.deleter
    def cauh(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def caflt(self):
        """
        CAFLT
        VT_VECTOR  |  VT_R4
        """
        return self._caflt

    @caflt.setter
    def caflt(self, value):
        _PropVariantClear(self)
        self._caflt = value
        self.vt = VT_VECTOR | VT_R4

    @caflt.deleter
    def caflt(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def cadbl(self):
        """
        CADBL
        VT_VECTOR  |  VT_R8
        """
        return self._cadbl

    @cadbl.setter
    def cadbl(self, value):
        _PropVariantClear(self)
        self._cadbl = value
        self.vt = VT_VECTOR | VT_R8

    @cadbl.deleter
    def cadbl(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def cabool(self):
        """
        CABOOL
        VT_VECTOR  |  VT_BOOL
        """
        return self._cabool

    @cabool.setter
    def cabool(self, value):
        _PropVariantClear(self)
        self._cabool = value
        self.vt = VT_VECTOR | VT_BOOL

    @cabool.deleter
    def cabool(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def cascode(self):
        """
        CASCODE
        VT_VECTOR  |  VT_ERROR
        """
        return self._cascode

    @cascode.setter
    def cascode(self, value):
        _PropVariantClear(self)
        self._cascode = value
        self.vt = VT_VECTOR | VT_ERROR

    @cascode.deleter
    def cascode(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def cacy(self):
        """
        CACY
        VT_VECTOR  |  VT_CY
        """
        return self._cacy

    @cacy.setter
    def cacy(self, value):
        _PropVariantClear(self)
        self._cacy = value
        self.vt = VT_VECTOR | VT_CY

    @cacy.deleter
    def cacy(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def cadate(self):
        """
        CADATE
        VT_VECTOR  |  VT_DATE
        """
        return self._cadate

    @cadate.setter
    def cadate(self, value):
        _PropVariantClear(self)
        self._cadate = value
        self.vt = VT_VECTOR | VT_DATE

    @cadate.deleter
    def cadate(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def cafiletime(self):
        """
        CAFILETIME
        VT_VECTOR  |  VT_FILETIME
        """
        return self._cafiletime

    @cafiletime.setter
    def cafiletime(self, value):
        _PropVariantClear(self)
        self._cafiletime = value
        self.vt = VT_VECTOR | VT_FILETIME

    @cafiletime.deleter
    def cafiletime(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def cauuid(self):
        """
        CACLSID
        VT_VECTOR  |  VT_CLSID
        """
        return self._cauuid

    @cauuid.setter
    def cauuid(self, value):
        _PropVariantClear(self)
        self._cauuid = value
        self.vt = VT_VECTOR | VT_CLSID

    @cauuid.deleter
    def cauuid(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def caclipdata(self):
        """
        CACLIPDATA
        VT_VECTOR  |  VT_CF
        """
        return self._caclipdata

    @caclipdata.setter
    def caclipdata(self, value):
        _PropVariantClear(self)
        self._caclipdata = value
        self.vt = VT_VECTOR | VT_CF

    @caclipdata.deleter
    def caclipdata(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def cabstr(self):
        """
        CABSTR
        VT_VECTOR  |  VT_BSTR
        """
        return self._cabstr

    @cabstr.setter
    def cabstr(self, value):
        _PropVariantClear(self)
        self._cabstr = value
        self.vt = VT_VECTOR | VT_BSTR

    @cabstr.deleter
    def cabstr(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def cabstrblob(self):
        """
        CABSTRBLOB
        VT_VECTOR  |  VT_BSTR
        """
        return self._cabstrblob

    @cabstrblob.setter
    def cabstrblob(self, value):
        _PropVariantClear(self)
        self._cabstrblob = value
        self.vt = VT_VECTOR | VT_BSTR

    @cabstrblob.deleter
    def cabstrblob(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def calpstr(self):
        """
        CALPSTR
        VT_VECTOR  |  VT_LPSTR
        """
        return self._calpstr

    @calpstr.setter
    def calpstr(self, value):
        _PropVariantClear(self)
        self._calpstr = value
        self.vt = VT_VECTOR | VT_LPSTR

    @calpstr.deleter
    def calpstr(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def calpwstr(self):
        """
        CALPWSTR
        VT_VECTOR  |  VT_LPWSTR
        """
        return self._calpwstr

    @calpwstr.setter
    def calpwstr(self, value):
        _PropVariantClear(self)
        self._calpwstr = value
        self.vt = VT_VECTOR | VT_LPWSTR

    @calpwstr.deleter
    def calpwstr(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def capropvar(self):
        """
        CAPROPVARIANT
        VT_VECTOR  |  VT_VARIANT
        """
        return self._capropvar

    @capropvar.setter
    def capropvar(self, value):
        _PropVariantClear(self)
        self._capropvar = value
        self.vt = VT_VECTOR | VT_VARIANT

    @capropvar.deleter
    def capropvar(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pcVal(self):
        """
        POINTER(c_char
        VT_BYREF | VT_I1
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        return variant.value

    @pcVal.setter
    def pcVal(self, value):
        _PropVariantClear(self)

        ctypes.cast(self._pcVal, POINTER(CHAR))[0] = value
        self.vt = VT_BYREF | VT_I1

    @pcVal.deleter
    def pcVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pbVal(self):
        """
        POINTER(c_byte)
        VT_BYREF | VT_UI1
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        return variant.value

    @pbVal.setter
    def pbVal(self, value):
        _PropVariantClear(self)
        ctypes.cast(self._pbVal, POINTER(BYTE))[0] = value
        self.vt = VT_BYREF | VT_UI1

    @pbVal.deleter
    def pbVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def piVal(self):
        """
        POINTER(c_short)
        VT_BYREF | VT_I2
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        return variant.value

    @piVal.setter
    def piVal(self, value):
        _PropVariantClear(self)

        ctypes.cast(self._piVal, POINTER(SHORT))[0] = value
        self.vt = VT_BYREF | VT_I2

    @piVal.deleter
    def piVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def puiVal(self):
        """
        POINTER(c_ushort)
        VT_BYREF | VT_UI2
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        return variant.value

    @puiVal.setter
    def puiVal(self, value):
        _PropVariantClear(self)

        ctypes.cast(self._puiVal, POINTER(USHORT))[0] = value
        self.vt = VT_BYREF | VT_UI2

    @puiVal.deleter
    def puiVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def plVal(self):
        """
        POINTER(c_long)
        VT_BYREF | VT_I4
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        return variant.value

    @plVal.setter
    def plVal(self, value):
        _PropVariantClear(self)

        ctypes.cast(self._plVal, POINTER(LONG))[0] = value
        self.vt = VT_BYREF | VT_I4

    @plVal.deleter
    def plVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pulVal(self):
        """
        POINTER(c_ulong)
        VT_BYREF | VT_UI4
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        return variant.value

    @pulVal.setter
    def pulVal(self, value):
        _PropVariantClear(self)

        ctypes.cast(self._pulVal, POINTER(ULONG))[0] = value
        self.vt = VT_BYREF | VT_UI4

    @pulVal.deleter
    def pulVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pintVal(self):
        """
        POINTER(c_int)
        VT_BYREF | VT_INT
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        return variant.value

    @pintVal.setter
    def pintVal(self, value):
        _PropVariantClear(self)

        ctypes.cast(self._pintVal, POINTER(INT))[0] = value
        self.vt = VT_BYREF | VT_INT

    @pintVal.deleter
    def pintVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def puintVal(self):
        """
        POINTER(c_uint)
        VT_BYREF | VT_UINT
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        return variant.value

    @puintVal.setter
    def puintVal(self, value):
        _PropVariantClear(self)

        ctypes.cast(self._puintVal, POINTER(UINT))[0] = value
        self.vt = VT_BYREF | VT_UINT

    @puintVal.deleter
    def puintVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pfltVal(self):
        """
        POINTER(c_float)
        VT_BYREF | VT_R4
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        return variant.value

    @pfltVal.setter
    def pfltVal(self, value):
        _PropVariantClear(self)

        ctypes.cast(self._pfltVal, POINTER(FLOAT))[0] = value
        self.vt = VT_BYREF | VT_R4

    @pfltVal.deleter
    def pfltVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pdblVal(self):
        """
        POINTER(c_double)
        VT_BYREF | VT_R8
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        return variant.value

    @pdblVal.setter
    def pdblVal(self, value):
        _PropVariantClear(self)

        ctypes.cast(self._pdblVal, POINTER(DOUBLE))[0] = value
        self.vt = VT_BYREF | VT_R8

    @pdblVal.deleter
    def pdblVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pboolVal(self):
        """
        POINTER(VARIANT_BOOL)
        VT_BYREF | VT_BOOL
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        return variant.value

    @pboolVal.setter
    def pboolVal(self, value):
        _PropVariantClear(self)

        ctypes.cast(self._pboolVal, POINTER(VARIANT_BOOL))[0] = value
        self.vt = VT_BYREF | VT_BOOL

    @pboolVal.deleter
    def pboolVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pdecVal(self):
        """
        POINTER(DECIMAL)
        VT_BYREF | VT_DECIMAL
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        return variant.value

    @pdecVal.setter
    def pdecVal(self, value):
        _PropVariantClear(self)

        ctypes.cast(self._pdecVal, POINTER(DECIMAL))[0] = value
        self.vt = VT_BYREF | VT_DECIMAL

    @pdecVal.deleter
    def pdecVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pscode(self):
        """
        POINTER(SCODE)
        VT_BYREF | VT_ERROR
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        return variant.value

    @pscode.setter
    def pscode(self, value):
        _PropVariantClear(self)

        ctypes.cast(self._pscode, POINTER(SCODE))[0] = value
        self.vt = VT_BYREF | VT_ERROR

    @pscode.deleter
    def pscode(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pcyVal(self):
        """
        POINTER(CY)
        VT_BYREF | VT_CY
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        return variant.value

    @pcyVal.setter
    def pcyVal(self, value):
        _PropVariantClear(self)

        ctypes.cast(self._pcyVal, POINTER(CY))[0] = value
        self.vt = VT_BYREF | VT_CY

    @pcyVal.deleter
    def pcyVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pdate(self):
        """
        POINTER(DATE)
        VT_BYREF | VT_DATE
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        return variant.value

    @pdate.setter
    def pdate(self, value):
        _PropVariantClear(self)

        ctypes.cast(self._pdate, POINTER(DOUBLE))[0] = value
        self.vt = VT_BYREF | VT_DATE

    @pdate.deleter
    def pdate(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pbstrVal(self):
        """
        POINTER(BSTR)
        VT_BYREF | VT_BSTR
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        return variant.value

    @pbstrVal.setter
    def pbstrVal(self, value):
        _PropVariantClear(self)

        ctypes.cast(self._pbstrVal, POINTER(BSTR))[0] = value
        self.vt = VT_BYREF | VT_BSTR

    @pbstrVal.deleter
    def pbstrVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def ppunkVal(self):
        """
        POINTER(POINTER(IUnknown))
        VT_BYREF | VT_UNKNOWN
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        return variant.value

    @ppunkVal.setter
    def ppunkVal(self, value):
        _PropVariantClear(self)

        ctypes.cast(self._ppunkVal, POINTER(POINTER(IUnknown)))[0] = value
        self.vt = VT_BYREF | VT_UNKNOWN

    @ppunkVal.deleter
    def ppunkVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def ppdispVal(self):
        """
        POINTER(POINTER(IDispatch))
        VT_BYREF | VT_DISPATCH
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        return variant.value

    @ppdispVal.setter
    def ppdispVal(self, value):
        _PropVariantClear(self)

        ctypes.cast(self._ppdispVal, POINTER(POINTER(IDispatch)))[0] = value
        self.vt = VT_BYREF | VT_DISPATCH

    @ppdispVal.deleter
    def ppdispVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pparray(self):
        """
        POINTER(POINTER(_safearray.SAFEARRAY))
        VT_BYREF | VT_ARRAY
        """
        variant = PROPVARIANT()
        _PropVariantCopy(variant, self)
        typ = _vartype_to_ctype[self.vt & ~VT_ARRAY]
        return ctypes.cast(variant._pparray, _midlSAFEARRAY(typ)).unpack()  # NOQA

    @pparray.setter
    def pparray(self, value):
        _PropVariantClear(self)
        ctypes.cast(
            self._pparray,
            POINTER(POINTER(_safearray.SAFEARRAY))
        )[0] = value
        self.vt = VT_BYREF | VT_ARRAY

    @pparray.deleter
    def pparray(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def pvarVal(self):
        """
        POINTER(VARIANT)
        VT_BYREF | VT_VARIANT
        """
        # apparently VariantCopyInd doesn't work always with
        # VT_BYREF|VT_VARIANT, so do it manually.
        return ctypes.cast(self._pvarVal, POINTER(PROPVARIANT))[0].value

    @pvarVal.setter
    def pvarVal(self, value):
        _PropVariantClear(self)

        ctypes.cast(self._pvarVal, POINTER(PROPVARIANT))[0] = value
        self.vt = VT_BYREF | VT_VARIANT

    @pvarVal.deleter
    def pvarVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    @property
    def decVal(self):
        """
        DECIMAL
        VT_DECIMAL
        """
        return self._decVal.as_decimal()

    @decVal.setter
    def decVal(self, value):
        _PropVariantClear(self)
        if decimal is not None and isinstance(value, decimal.Decimal):
            self._decVal = DECIMAL(int(round(value * 10000)))
            self.vt = VT_DECIMAL

    @decVal.deleter
    def decVal(self):
        self.vt = VT_NULL
        _PropVariantClear(self)

    def __init__(self, *args):
        super(tagPROPVARIANT, self).__init__()

        if args:
            self.value = args[0]

    def __del__(self):
        if self._b_needsfree_:
            # XXX This does not work.  _b_needsfree_ is never
            # set because the buffer is internal to the object.
            _PropVariantClear(self)

    def __repr__(self):
        if self.vt & VT_BYREF:
            return "PROPVARIANT(vt=0x%x, ctypes.byref(%r))" % (self.vt, self[0])
        return "PROPVARIANT(vt=0x%x, %r)" % (self.vt, self.value)

    def from_param(cls, value):
        if isinstance(value, cls):
            return value
        return cls(value)

    from_param = classmethod(from_param)

    def __setitem__(self, index, value):
        # This method allows to change the value of a
        # (VT_BYREF|VT_xxx) variant in place.

        if index != 0:
            raise IndexError(index)

        if not self.vt & VT_BYREF:
            raise TypeError(
                "set_ctypes.byref requires a VT_BYREF PROPVARIANT instance")

        vt = self.vt & ~VT_BYREF

        if isinstance(value, _byref_type):
            ref = value._obj  # NOQA
            self.__keepref = value
            value = ref
        elif isinstance(value, _Pointer):
            ref = value.contents
            self.__keepref = value
            value = ref

        if vt == VT_UI1:
            self.pbVal = value
        elif vt == VT_I2:
            self.piVal = value
        elif vt == VT_I4:
            self.plVal = value
        elif vt == VT_I8:
            self.hval = value
        elif vt == VT_R4:
            self.pfltVal = value
        elif vt == VT_R8:
            self.pdblVal = value
        elif vt == VT_BOOL:
            self.pboolVal = value
        elif vt == VT_ERROR:
            self.pscode = value
        elif vt == VT_CY:
            self.pcyVal = value
        elif vt == VT_DATE:
            self.pdate = value
        elif vt == VT_BSTR:
            self.pbstrVal = value
        elif vt == VT_UNKNOWN:
            self.ppunkVal = value
        elif vt == VT_DISPATCH:
            self.ppdispVal = value
        elif vt == VT_ARRAY:
            self.pparray = value
        elif vt == VT_VARIANT:
            self.pvarVal = value
        elif vt == VT_DECIMAL:
            self.pdecVal = value
        elif vt == VT_I1:
            self.pcVal = value
        elif vt == VT_UI2:
            self.puiVal = value
        elif vt == VT_UI4:
            self.pulVal = value
        elif vt == VT_UI8:
            self.uhVal = value
        elif vt == VT_INT:
            self.pintVal = value
        elif vt == VT_UINT:
            self.puintVal = value
        else:
            self.ctypes.byref = value

    # see also c:/sf/pywin32/com/win32com/src/oleargs.cpp 54
    def _set_value(self, value):
        if value is None:
            _PropVariantClear(self)
            self.vt = VT_NULL
        elif self.vt:
            vt = self.vt
            if vt & VT_BYREF:
                self[0] = value

            elif vt == VT_VECTOR | VT_UI1:
                self.caub = value
            elif vt == VT_BLOB:
                self.blob = value
            elif vt == VT_LPSTR:
                self.pszVal = value
            elif vt == VT_LPWSTR:
                self.pwszVal = value
            elif vt == VT_UI1:
                self.bVal = value
            elif vt == VT_I2:
                self.iVal = value
            elif vt == VT_I4:
                self.lVal = value
            elif vt == VT_I8:
                self.hval = value
            elif vt == VT_R4:
                self.fltVal = value
            elif vt == VT_R8:
                self.dblVal = value
            elif vt == VT_BOOL:
                self.boolVal = value
            elif vt == VT_ERROR:
                self.scode = value
            elif vt == VT_CY:
                self.cyVal = value
            elif vt == VT_DATE:
                self.date = value
            elif vt == VT_BSTR:
                self.bstrVal = value
            elif vt == VT_UNKNOWN:
                self.punkVal = value
            elif vt == VT_DISPATCH:
                self.pdispVal = value
            elif vt == VT_ARRAY:
                self.parray = value
            elif vt == VT_VARIANT:
                self.varVal = value
            elif vt == VT_DECIMAL:
                self.decVal = value
            elif vt == VT_I1:
                self.cVal = value
            elif vt == VT_UI2:
                self.uiVal = value
            elif vt == VT_UI4:
                self.ulVal = value
            elif vt == VT_UI8:
                self.uhval = value
            elif vt == VT_INT:
                self.intVal = value
            elif vt == VT_UINT:
                self.uintVal = value
            else:
                raise TypeError("Unknown VT_* type ({0})".format(vt))

        elif (
                hasattr(value, '__len__') and
                len(value) == 0 and
                not isinstance(value, str)
        ):
            _PropVariantClear(self)
            self.vt = VT_NULL

        # since bool is a subclass of int, this check must come before
        # the check for int
        elif isinstance(value, bool):
            self.boolVal = value
        elif isinstance(value, INT):
            self.lVal = value
        elif isinstance(value, int):
            if value < 0:
                self.cVal = value
                if self.cVal == value:
                    return
                self.iVal = value
                if self.iVal == value:
                    return
                self.lVal = value
                if self.lVal == value:
                    return
                self.llVal = value
                if self.llVal == value:
                    return
                self.intVal = value
                if self.intVal == value:
                    return
            else:
                self.bVal = value
                if self.bVal == value:
                    return
                self.uiVal = value
                if self.uiVal == value:
                    return
                self.ulVal = value
                if self.ulVal == value:
                    return
                self.ullVal = value
                if self.ullVal == value:
                    return
                self.uintVal = value
                if self.uintVal == value:
                    return

            # VT_R8 is last resort.
            self.dblVal = float(value)

        elif isinstance(value, (UBYTE, CHAR)):
            self.bVal = value  # VT_UI1
        elif isinstance(value, (BYTE, CHAR)):
            self.cVal = value  # VT_I1
        elif isinstance(value, FLOAT):
            self.fltVal = value  # VT_R4
        elif isinstance(value, DOUBLE):
            self.dblVal = value  # VT_R8
        elif isinstance(value, VARIANT_BOOL):
            self.boolVal = value  # VT_BOOL
        elif isinstance(value, BSTR):
            self.bstrVal = value  # VT_BSTR
        elif isinstance(value, SHORT):
            self.iVal = value  # VT_I2
        elif isinstance(value, USHORT):
            self.uiVal = value  # VT_UI2
        elif isinstance(value, LONG):
            self.lVal = value  # VT_I4
        elif isinstance(value, ULONG):
            self.ulVal = value  # VT_UI4
        elif isinstance(value, INT):
            self.intVal = value  # VT_INT
        elif isinstance(value, UINT):
            self.uintVal = value  # VT_UINT
        elif isinstance(value, (INT64, LONGLONG)):
            self.hval = value  # VT_I8
        elif isinstance(value, (UINT64, ULONGLONG)):
            self.uhval = value  # VT_UI8
        elif isinstance(value, float):
            self.fltVal = value
            if self.fltVal == value:
                return

            self.dblVal = value
        elif isinstance(value, str):
            self.bstrVal = value
        elif isinstance(value, datetime.datetime):
            self.date = value
        elif npsupport.isdatetime64(value):
            self.date = value
        elif decimal is not None and isinstance(value, decimal.Decimal):
            self.decVal = value
        elif (
                isinstance(value, POINTER(IDispatch)) or
                isinstance(getattr(value, "_comobj", None), POINTER(IDispatch))
        ):
            self.pdispVal = value
        elif isinstance(value, POINTER(IUnknown)):
            self.punkVal = value
        elif (
                isinstance(value, (list, tuple, array.array)) or
                npsupport.isndarray(value)
        ):
            self.parray = value
        elif (
            isinstance(value, ctypes.Structure) and
            hasattr(value, "_recordinfo_")
        ):
            self.pRecInfo = value
        elif isinstance(value, PROPVARIANT):
            _PropVariantCopy(self, value)
        elif isinstance(value, _byref_type):
            ref = value._obj  # NOQA
            self.vt = _ctype_to_vartype[type(ref)] | VT_BYREF
            self[0] = value
        elif isinstance(value, _Pointer):
            ref = value.contents
            self.vt = _ctype_to_vartype[type(ref)] | VT_BYREF
            self[0] = value

        else:
            raise TypeError("Cannot put %r in PROPVARIANT" % value)
        # buffer ->  SAFEARRAY of VT_UI1 ?

    # c:/sf/pywin32/com/win32com/src/oleargs.cpp 197
    def _get_value(self, dynamic=False):  # NOQA
        vt = self.vt

        if vt in (VT_EMPTY, VT_NULL):
            return None
        if vt & VT_BYREF:
            return self[0]
        if vt == VT_VECTOR | VT_I1:
            return self.cab
        if vt == VT_VECTOR | VT_UI1:
            return self.caub
        if vt == VT_VECTOR | VT_LPWSTR:
            return self.calpwstr
        if vt == VT_BLOB:
            return self.blob
        if vt == VT_LPSTR:
            return self.pszVal
        if vt == VT_LPWSTR:
            return self.pwszVal
        if vt == VT_UI1:
            return self.bVal
        if vt == VT_I2:
            return self.iVal
        if vt == VT_I4:
            return self.lVal
        if vt == VT_I8:
            return self.hVal.value
        if vt == VT_R4:
            return self.fltVal
        if vt == VT_R8:
            return self.dblVal
        if vt == VT_BOOL:
            return self.boolVal
        if vt == VT_ERROR:
            return self.scode
        if vt == VT_CY:
            return self.cyVal
        if vt == VT_DATE:
            return self.date
        if vt == VT_BSTR:
            return self.bstrVal
        if vt == VT_UNKNOWN:
            return self.punkVal
        if vt == VT_DISPATCH:
            return self.pdispVal
        if vt == VT_ARRAY:
            return self.parray
        if vt == VT_VARIANT:
            return self.varVal
        if vt == VT_DECIMAL:
            return self.decVal
        if vt == VT_I1:
            return self.cVal
        if vt == VT_UI2:
            return self.uiVal
        if vt == VT_UI4:
            return self.ulVal
        if vt == VT_UI8:
            # variant = ctypes.cast(ctypes.byref(self), POINTER(VARIANT)).contents.Union_2
            # print(dir(variant))
            return self.uhVal.value
        if vt == VT_INT:
            return self.intVal
        if vt == VT_UINT:
            return self.uintVal
        if vt == VT_FILETIME:
            return self.filetime.value
        if vt == VT_CLSID:
            return self.puuid.contents

        raise NotImplementedError("typecode %d = 0x%x)" % (vt, vt))

    def __getitem__(self, index):
        if index != 0:
            raise IndexError(index)

        if not self.vt & VT_BYREF:
            raise TypeError(
                "set_ctypes.byref requires a VT_BYREF PROPVARIANT instance")

        vt = self.vt & ~VT_BYREF

        if vt == VT_UI1:
            return self.pbVal
        if vt == VT_I2:
            return self.piVal
        if vt == VT_I4:
            return self.plVal
        if vt == VT_I8:
            return self.pllVal
        if vt == VT_R4:
            return self.pfltVal
        if vt == VT_R8:
            return self.pdblVal
        if vt == VT_BOOL:
            return self.pboolVal
        if vt == VT_ERROR:
            return self.pscode
        if vt == VT_CY:
            return self.pcyVal
        if vt == VT_DATE:
            return self.pdate
        if vt == VT_BSTR:
            return self.pbstrVal
        if vt == VT_UNKNOWN:
            return self.ppunkVal
        if vt == VT_DISPATCH:
            return self.ppdispVal
        if vt == VT_ARRAY:
            return self.pparray
        if vt == VT_VARIANT:
            return self.pvarVal
        if vt == VT_DECIMAL:
            return self.pdecVal
        if vt == VT_I1:
            return self.pcVal
        if vt == VT_UI2:
            return self.puiVal
        if vt == VT_UI4:
            return self.pulVal
        if vt == VT_UI8:
            return self.pullVal
        if vt == VT_INT:
            return self.pintVal
        if vt == VT_UINT:
            return self.puintVal

        return self.ctypes.byref

    value = property(_get_value, _set_value)

    def __ctypes_from_outparam__(self):
        # XXX Manual resource management, because of the PROPVARIANT bug:
        result = self.value
        self.value = None
        return result

    def ChangeType(self, typecode):
        _PropVariantChangeType(self, self, 0, typecode)


PROPVARIANT = tagPROPVARIANT
PPROPVARIANT = POINTER(tagPROPVARIANT)


class tagPROPSPEC(ctypes.Structure):
    class DUMMYUNIONNAME(ctypes.Union):
        _fields_ = [
            ('propid', PROPID),
            ('lpwstr', LPOLESTR),
        ]

    _fields_ = [
        ('ulKind', ULONG),
        ('DUMMYUNIONNAME', DUMMYUNIONNAME),
    ]
    _anonymous_ = ('DUMMYUNIONNAME',)


PROPSPEC = tagPROPSPEC


class tagSTATPROPSTG(ctypes.Structure):
    _fields_ = [
        ('lpwstrName', LPOLESTR),
        ('propid', PROPID),
        ('vt', VARTYPE),
    ]


STATPROPSTG = tagSTATPROPSTG


class tagSTATPROPSETSTG(ctypes.Structure):
    _fields_ = [
        ('fmtid', FMTID),
        ('clsid', CLSID),
        ('grfFlags', DWORD),
        ('mtime', FILETIME),
        ('ctime', FILETIME),
        ('atime', FILETIME),
        ('dwOSVersion', DWORD),
    ]


STATPROPSETSTG = tagSTATPROPSETSTG


class tagSERIALIZEDPROPERTYVALUE(ctypes.Structure):
    _fields_ = [
        ('dwType', DWORD),
        ('rgb', BYTE * 1),
    ]


SERIALIZEDPROPERTYVALUE = tagSERIALIZEDPROPERTYVALUE


class tagSTATSTG(ctypes.Structure):
    _fields_ = [
        ('pwcsName', LPOLESTR),
        ('type', DWORD),
        ('cbSize', ULARGE_INTEGER),
        ('mtime', FILETIME),
        ('ctime', FILETIME),
        ('atime', FILETIME),
        ('grfMode', DWORD),
        ('grfLocksSupported', DWORD),
        ('clsid', CLSID),
        ('grfStateBits', DWORD),
        ('reserved', DWORD),
    ]


STATSTG = tagSTATSTG


class Vector(ctypes.Structure):

    def __init__(self, *args):
        if len(args) == 1:
            val = args[1]
            super().__init__()
            self.value = val
        else:
            super().__init__(*args)

    @property
    def value(self):
        if isinstance(self, (tagBLOB, tagBSTRBLOB)):
            size = self.cbSize
            if isinstance(self, tagBLOB):
                from .mmreg import WAVEFORMATEXTENSIBLE, WAVEFORMATEX, WAVE_FORMAT_EXTENSIBLE

                w_format = ctypes.cast(self.pBlobData, POINTER(WAVEFORMATEX)).contents
                # PROPERTYKEY({4B361010-DEF7-43A1-A5DC-071D955B62F7}, 12) = swap center and subwoofer outputs
                # PROPERTYKEY({4B361010-DEF7-43A1-A5DC-071D955B62F7}, 9) = bass management
                # PROPERTYKEY({E0A941A0-88A2-4DF5-8D6B-DD20BB06E8FB}, 1) = stereo upmix
                # PROPERTYKEY({4B361010-DEF7-43A1-A5DC-071D955B62F7}, 14) = time alignment on/off
                # PROPERTYKEY({4B361010-DEF7-43A1-A5DC-071D955B62F7}, 15) = channel db
                # PROPERTYKEY({4B361010-DEF7-43A1-A5DC-071D955B62F7}, 16) = channel distance

                if w_format.wFormatTag == WAVE_FORMAT_EXTENSIBLE:
                    return ctypes.cast(self.pBlobData, POINTER(WAVEFORMATEXTENSIBLE)).contents

                if (
                    w_format.nChannels <= 2 and
                    16 >= w_format.wBitsPerSample > 0 and
                    w_format.wBitsPerSample % 8 == 0
                ):
                    return w_format

                data = self.pBlobData

            else:
                data = self.pData
        else:
            size = self.cElems // ctypes.sizeof(self.pElems._type_)
            data = self.pElems

        res = []

        for i in range(size):
            res.append(data[i])

        return res

    @value.setter
    def value(self, val):
        if isinstance(self, (tagBLOB, tagBSTRBLOB)):
            if isinstance(self, tagBLOB):
                data_type = self.pBlobData._type_
            else:
                data_type = self.pData._type_

            self.cbSize = len(val)

        else:
            data_type = self.pElems._type_
            self.cElems = len(val) * ctypes.sizeof(data_type)

        data = (data_type * len(val))()

        for i, v in enumerate(val):
            data[i] = val

        if isinstance(self, tagBLOB):
            self.pBlobData = data

        elif isinstance(self, tagBSTRBLOB):
            self.pData = data

        else:
            self.pElems = data

    def __add__(self, other):
        if isinstance(other, self.__class__):
            other = other.value

        if not isinstance(other, list):
            raise ValueError(
                'you can only add a list '
                'or {0} instance together'.format(self.__class__.__name__)
            )

        value = self.value
        value += other
        return self.__class__(value)

    def __iadd__(self, other):
        if isinstance(other, self.__class__):
            other = other.value

        if not isinstance(other, list):
            raise ValueError(
                'you can only add a list '
                'or {0} instance together'.format(self.__class__.__name__)
            )

        value = self.value
        value += other
        self.value = value
        return self

    def __imul__(self, other):
        res = []
        value = self.value
        for _ in range(other):
            res.extend(value)

        self.value = res
        return self

    def __mul__(self, other):
        res = []
        value = self.value
        for _ in range(other):
            res.extend(value)

        return self.__class__(res)

    def __rmul__(self, other):
        res = []
        value = self.value
        for _ in range(other):
            res.extend(value)

        return self.__class__(res)

    def __contains__(self, item):
        value = self.value

        return item in value

    def __delitem__(self, key):
        value = self.value

        del value[key]

        self.value = value

    def __len__(self):
        value = self.value

        return len(value)

    def __reversed__(self):
        value = self.value
        value.reverse()

        self.value = value
        return self

    def append(self, obj):
        value = self.value
        value.append(obj)
        self.value = value

    def clear(self):
        self.value = []

    def copy(self):
        value = self.value
        return self.__class__(value)

    def count(self, val):
        value = self.value
        return value.count(val)

    def extend(self, iterable):
        if isinstance(iterable, self.__class__):
            iterable = iterable.value

        if not isinstance(iterable, list):
            raise ValueError(
                'you can only extend with a list '
                'or {0} instance together'.format(self.__class__.__name__)
            )

        value = self.value
        value.extend(iterable)
        self.value = value

    def index(self, val, start=None, stop=None):
        value = self.value

        if start is None:
            start = 0

        if stop is None:
            stop = len(self.value)

        return value.index(val, start, stop)

    def insert(self, index, obj):
        pass

    def pop(self, index=None):
        value = self.value

        if index is None:
            index = len(value) - 1

        res = value.pop(index)
        self.value = value

        return res

    def remove(self, val):
        value = self.value

        if val not in value:
            raise ValueError(val)

        value.remove(val)
        self.value = value

    def reverse(self):
        value = self.value
        value.reverse()
        self.value = value

    def sort(self, *, key=None, reverse=None):
        value = self.value

        if key is None and reverse is None:
            value.sort()

        elif key is None:
            value.sort(reverse=reverse)

        elif reverse is None:
            value.sort(key=key)

        else:
            value.sort(key=key, reverse=reverse)

        self.value = value

    def __getitem__(self, item):
        value = self.value
        try:
            return value[item]
        except IndexError:
            raise IndexError(item)

    def __setitem__(self, key, value):
        val = self.value

        if isinstance(key, int):
            if key >= len(val):
                raise IndexError(key)

            if isinstance(self, tagBLOB):
                self.pBlobData[key] = value
            elif isinstance(self, tagBSTRBLOB):
                self.pData[key] = value
            else:
                self.pElems[key] = value

        elif isinstance(key, slice):
            start = key.start
            stop = key.stop

            if start is None:
                start = 0
            if stop is None:
                stop = len(value)

            count = stop - start
            if len(value) < count:
                raise IndexError(key)

            if stop > len(val):
                raise IndexError(key)
            if start > len(val):
                raise IndexError

            for i, item in enumerate(value):
                if isinstance(self, tagBLOB):
                    self.pBlobData[i + start] = item
                elif isinstance(self, tagBSTRBLOB):
                    self.pData[i + start] = item
                else:
                    self.pElems[i + start] = item

    def __iter__(self):
        for item in self.value:
            yield item

    def __str__(self):
        return str(self.value)


class tagBLOB(Vector):

    _fields_ = [
        ('cbSize', ULONG),
        ('pBlobData', POINTER(UBYTE)),
    ]


BLOB = tagBLOB


class tagBSTRBLOB(Vector):
    _fields_ = [
        ('cbSize', ULONG),
        ('pData', POINTER(BYTE))
    ]


BSTRBLOB = tagBSTRBLOB


class tagCLIPDATA(ctypes.Structure):
    _fields_ = [
        ('cbSize', ULONG),
        ('ulClipFmt', LONG),
        ('pClipData', POINTER(BYTE))
    ]


CLIPDATA = tagCLIPDATA


class tagCAC(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(CHAR)),
    ]


CAC = tagCAC


class tagCAUB(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(UBYTE)),
    ]


CAUB = tagCAUB


class tagCAI(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(SHORT)),
    ]


CAI = tagCAI


class tagCAUI(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(USHORT)),
    ]


CAUI = tagCAUI


class tagCAL(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(LONG)),
    ]


CAL = tagCAL


class tagCAUL(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(ULONG)),
    ]


CAUL = tagCAUL


class tagCAFLT(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(FLOAT)),
    ]


CAFLT = tagCAFLT


class tagCADBL(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(DOUBLE)),
    ]


CADBL = tagCADBL


class tagCACY(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(CY)),
    ]


CACY = tagCACY


class tagCADATE(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(DOUBLE)),
    ]


CADATE = tagCADATE


class tagCABSTR(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(BSTR)),
    ]


CABSTR = tagCABSTR


class tagCABSTRBLOB(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(BSTRBLOB)),
    ]


CABSTRBLOB = tagCABSTRBLOB


class tagCABOOL(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(VARIANT_BOOL)),
    ]


CABOOL = tagCABOOL


class tagCASCODE(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(SCODE)),
    ]


CASCODE = tagCASCODE


class tagCAPROPVARIANT(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(PROPVARIANT)),
    ]


CAPROPVARIANT = tagCAPROPVARIANT


class tagCAH(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(LARGE_INTEGER)),
    ]


CAH = tagCAH


class tagCAUH(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(ULARGE_INTEGER)),
    ]


CAUH = tagCAUH


class tagCALPSTR(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(LPSTR)),
    ]


CALPSTR = tagCALPSTR


class tagCALPWSTR(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(LPWSTR)),
    ]


CALPWSTR = tagCALPWSTR


class tagCAFILETIME(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(FILETIME)),
    ]


CAFILETIME = tagCAFILETIME


class tagCACLIPDATA(Vector):

    _fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(CLIPDATA)),
    ]


CACLIPDATA = tagCACLIPDATA


class tagCACLSID(Vector):

    fields_ = [
        ('cElems', ULONG),
        ('pElems', POINTER(CLSID)),
    ]


CACLSID = tagCACLSID

# =================  Interface Forward Declerations  =================
IID_IEnumSTATPROPSTG = GUID("{00000139-0000-0000-C000-000000000046}")


class IEnumSTATPROPSTG(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IEnumSTATPROPSTG
    _idlflags_ = []


IID_IEnumSTATPROPSETSTG = GUID("{0000013B-0000-0000-C000-000000000046}")


class IEnumSTATPROPSETSTG(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IEnumSTATPROPSETSTG
    _idlflags_ = []


IID_ISequentialStream = GUID("{0C733A30-2A1C-11CE-ADE5-00AA0044773D}")


class ISequentialStream(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_ISequentialStream
    _idlflags_ = []


IID_IStream = GUID("{0000000C-0000-0000-C000-000000000046}")


class IStream(ISequentialStream):
    _case_insensitive_ = True
    _idlflags_ = []
    _iid_ = IID_IStream


IID_IStorage = GUID("{0000000B-0000-0000-C000-000000000046}")


class IStorage(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IStorage
    _idlflags_ = []


IID_IEnumSTATSTG = GUID("{0000000D-0000-0000-C000-000000000046}")


class IEnumSTATSTG(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IEnumSTATSTG
    _idlflags_ = []


# ====================================================================

IID_IPropertyStorage = GUID("{00000138-0000-0000-C000-000000000046}")


class IPropertyStorage(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertyStorage
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Method ReadMultiple')],
            HRESULT,
            'ReadMultiple',
            (['in'], ULONG, 'cpspec'),
            (['in'], PROPSPEC * 0, 'rgpspec'),
            (['out'], PROPVARIANT * 0, 'rgpropvar'),
        ),
        COMMETHOD(
            [helpstring('Method WriteMultiple')],
            HRESULT,
            'WriteMultiple',
            (['in'], ULONG, 'cpspec'),
            (['in'], PROPSPEC * 0, 'rgpspec'),
            (['in'], PROPVARIANT * 0, 'rgpropvar'),
            (['in'], PROPID, 'propidNameFirst'),
        ),
        COMMETHOD(
            [helpstring('Method DeleteMultiple')],
            HRESULT,
            'DeleteMultiple',
            (['in'], ULONG, 'cpspec'),
            (['in'], PROPSPEC * 0, 'rgpspec'),
        ),
        COMMETHOD(
            [helpstring('Method ReadPropertyNames')],
            HRESULT,
            'ReadPropertyNames',
            (['in'], ULONG, 'cpropid'),
            (['in'], PROPID * 0, 'rgpropid'),
            (['out'], LPOLESTR * 0, 'rglpwstrName'),
        ),
        COMMETHOD(
            [helpstring('Method WritePropertyNames')],
            HRESULT,
            'WritePropertyNames',
            (['in'], ULONG, 'cpropid'),
            (['in'], PROPID * 0, 'rgpropid'),
            (['in'], LPOLESTR * 0, 'rglpwstrName'),
        ),
        COMMETHOD(
            [helpstring('Method DeletePropertyNames')],
            HRESULT,
            'DeletePropertyNames',
            (['in'], ULONG, 'cpropid'),
            (['in'], PROPID * 0, 'rgpropid'),
        ),
        COMMETHOD(
            [helpstring('Method Commit')],
            HRESULT,
            'Commit',
            (['in'], DWORD, 'grfCommitFlags'),
        ),
        COMMETHOD(
            [helpstring('Method Revert')],
            HRESULT,
            'Revert',
        ),
        COMMETHOD(
            [helpstring('Method Enum')],
            HRESULT,
            'Enum',
            (
                ['out'],
                POINTER(POINTER(IEnumSTATPROPSTG)),
                'ppenum'
            ),
        ),
        COMMETHOD(
            [helpstring('Method SetTimes')],
            HRESULT,
            'SetTimes',
            (['in'], POINTER(FILETIME), 'pctime'),
            (['in'], POINTER(FILETIME), 'patime'),
            (['in'], POINTER(FILETIME), 'pmtime'),
        ),
        COMMETHOD(
            [helpstring('Method SetClass')],
            HRESULT,
            'SetClass',
            (['in'], REFCLSID, 'clsid'),
        ),
        COMMETHOD(
            [helpstring('Method Stat')],
            HRESULT,
            'Stat',
            (['out'], POINTER(STATPROPSETSTG), 'pstatpsstg'),
        ),
    ]


IID_IPropertySetStorage = GUID("{0000013A-0000-0000-C000-000000000046}")


class IPropertySetStorage(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertySetStorage
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Method Create')],
            HRESULT,
            'Create',
            (['in'], REFFMTID, 'rfmtid'),
            (['unique', 'in'], POINTER(CLSID), 'pclsid'),
            (['in'], DWORD, 'grfFlags'),
            (['in'], DWORD, 'grfMode'),
            (
                ['out'],
                POINTER(POINTER(IPropertyStorage)),
                'ppprstg'
            ),
        ),
        COMMETHOD(
            [helpstring('Method Open')],
            HRESULT,
            'Open',
            (['in'], REFFMTID, 'rfmtid'),
            (['in'], DWORD, 'grfMode'),
            (
                ['out'],
                POINTER(POINTER(IPropertyStorage)),
                'ppprstg'
            ),
        ),
        COMMETHOD(
            [helpstring('Method Delete')],
            HRESULT,
            'Delete',
            (['in'], REFFMTID, 'rfmtid'),
        ),
        COMMETHOD(
            [helpstring('Method Enum')],
            HRESULT,
            'Enum',
            (
                ['out'],
                POINTER(POINTER(IEnumSTATPROPSETSTG)),
                'ppenum'
            ),
        ),
    ]


class tagVersionedStream(ctypes.Structure):
    _fields_ = [
        ('guidVersion', GUID),
        ('pStream', POINTER(IStream)),
    ]


VERSIONEDSTREAM = tagVersionedStream
LPVERSIONEDSTREAM = POINTER(tagVersionedStream)

# Flags for IPropertySetStorage::Create
PROPSETFLAG_DEFAULT = 0
PROPSETFLAG_NONSIMPLE = 1
PROPSETFLAG_ANSI = 2

# (This flag is only supported on StgCreatePropStg & StgOpenPropStg
PROPSETFLAG_UNBUFFERED = 4
# (This flag causes a version-1 property set to be created
PROPSETFLAG_CASE_SENSITIVE = 8
# Flags for the reserved PID_BEHAVIOR property
PROPSET_BEHAVIOR_CASE_SENSITIVE = 1

PROPVAR_PAD1 = WORD
PROPVAR_PAD2 = WORD
PROPVAR_PAD3 = WORD


# =======================  PROPVARIANT  =======================

class _Union_1(ctypes.Union):
    class tag_inner_PROPVARIANT(ctypes.Structure):
        class _Union_2(ctypes.Union):
            _fields_ = [
                ('cVal', CHAR),  # VT_I1
                ('bVal', UBYTE),  # VT_UI1
                ('iVal', SHORT),  # VT_I2
                ('uiVal', USHORT),  # VT_UI2
                ('lVal', LONG),  # VT_I4
                ('ulVal', ULONG),  # VT_UI4
                ('intVal', INT),  # VT_INT
                ('uintVal', UINT),  # VT_UINT
                ('hVal', LARGE_INTEGER),  # VT_I8
                ('uhVal', ULARGE_INTEGER),  # VT_UI8
                ('fltVal', FLOAT),  # VT_R4
                ('dblVal', DOUBLE),  # VT_R8
                ('boolVal', VARIANT_BOOL),  # VT_BOOL
                ('__OBSOLETE__VARIANT_BOOL', VARIANT_BOOL),  #
                ('scode', SCODE),  # VT_ERROR
                ('cyVal', CY),  # VT_CY
                ('date', DOUBLE),  # VT_DATE
                ('filetime', FILETIME),  # VT_FILETIME
                ('puuid', POINTER(CLSID)),  # VT_CLSID
                ('pclipdata', POINTER(CLIPDATA)),  # VT_CF
                ('bstrVal', BSTR),  # VT_BSTR
                ('bstrblobVal', BSTRBLOB),  # VT_BSTR_BLOB
                ('blob', BLOB),  # VT_BLOB, VT_BLOBOBJECT
                ('pszVal', STRING),  # VT_LPSTR
                ('pwszVal', WSTRING),  # VT_LPWSTR
                ('punkVal', POINTER(IUnknown)),  # VT_UNKNOWN
                ('pdispVal', POINTER(IDispatch)),  # VT_DISPATCH
                ('pStream', POINTER(IStream)),  # VT_STREAM, VT_STREAMED_OBJECT
                ('pStorage', POINTER(IStorage)),  # VT_STORAGE, VT_STORED_OBJECT
                ('pVersionedStream', LPVERSIONEDSTREAM),  # VT_VERSIONED_STREAM
                ('parray', LPSAFEARRAY),  # VT_ARRAY | VT_*
                ('cac', CAC),  # VT_VECTOR | VT_I1
                ('caub', CAUB),  # VT_VECTOR | VT_UI1
                ('cai', CAI),  # VT_VECTOR | VT_I2
                ('caui', CAUI),  # VT_VECTOR | VT_UI2
                ('cal', CAL),  # VT_VECTOR | VT_I4
                ('caul', CAUL),  # VT_VECTOR | VT_UI4
                ('cah', CAH),  # VT_VECTOR | VT_I8
                ('cauh', CAUH),  # VT_VECTOR | VT_UI8
                ('caflt', CAFLT),  # VT_VECTOR | VT_R4
                ('cadbl', CADBL),  # VT_VECTOR | VT_R8
                ('cabool', CABOOL),  # VT_VECTOR | VT_BOOL
                ('cascode', CASCODE),  # VT_VECTOR | VT_ERROR
                ('cacy', CACY),  # VT_VECTOR | VT_CY
                ('cadate', CADATE),  # VT_VECTOR | VT_DATE
                ('cafiletime', CAFILETIME),  # VT_VECTOR | VT_FILETIME
                ('cauuid', CACLSID),  # VT_VECTOR | VT_CLSID
                ('caclipdata', CACLIPDATA),  # VT_VECTOR | VT_CF
                ('cabstr', CABSTR),  # VT_VECTOR | VT_BSTR
                ('cabstrblob', CABSTRBLOB),  # VT_VECTOR | VT_BSTR
                ('calpstr', CALPSTR),  # VT_VECTOR | VT_LPSTR
                ('calpwstr', CALPWSTR),  # VT_VECTOR | VT_LPWSTR
                ('capropvar', CAPROPVARIANT),  # VT_VECTOR | VT_VARIANT
                ('pcVal', POINTER(CHAR)),  # VT_BYREF | VT_I1
                ('pbVal', POINTER(UBYTE)),  # VT_BYREF | VT_UI1
                ('piVal', POINTER(SHORT)),  # VT_BYREF | VT_I2
                ('puiVal', POINTER(USHORT)),  # VT_BYREF | VT_UI2
                ('plVal', POINTER(LONG)),  # VT_BYREF | VT_I4
                ('pulVal', POINTER(ULONG)),  # VT_BYREF | VT_UI4
                ('pintVal', POINTER(INT)),  # VT_BYREF | VT_INT
                ('puintVal', POINTER(UINT)),  # VT_BYREF | VT_UINT
                ('pfltVal', POINTER(FLOAT)),  # VT_BYREF | VT_R4
                ('pdblVal', POINTER(DOUBLE)),  # VT_BYREF | VT_R8
                ('pboolVal', POINTER(VARIANT_BOOL)),  # VT_BYREF | VT_BOOL
                ('pdecVal', POINTER(DECIMAL)),  # VT_BYREF | VT_DECIMAL
                ('pscode', POINTER(SCODE)),  # VT_BYREF | VT_ERROR
                ('pcyVal', POINTER(CY)),  # VT_BYREF | VT_CY
                ('pdate', POINTER(DOUBLE)),  # VT_BYREF | VT_DATE
                ('pbstrVal', POINTER(BSTR)),  # VT_BYREF | VT_BSTR
                ('ppunkVal', POINTER(POINTER(IUnknown))),
                # VT_BYREF | VT_UNKNOWN
                ('ppdispVal', POINTER(POINTER(IDispatch))),
                # VT_BYREF | VT_DISPATCH
                ('pparray', POINTER(LPSAFEARRAY)),  # VT_BYREF | VT_ARRAY | VT_*
                ('pvarVal', POINTER(PROPVARIANT)),  # VT_BYREF | VT_VARIANT
            ]

        _fields_ = [
            ('vt', VARTYPE),
            ('wReserved1', PROPVAR_PAD1),
            ('wReserved2', PROPVAR_PAD2),
            ('wReserved3', PROPVAR_PAD3),
            ('_Union_2', _Union_2),
        ]

        _anonymous_ = ('_Union_2',)

    _fields_ = [
        ('tag_inner_PROPVARIANT', tag_inner_PROPVARIANT),
        ('decVal', DECIMAL),
    ]

    _anonymous_ = ('tag_inner_PROPVARIANT',)


tagPROPVARIANT._anonymous_ = ('_Union_1',)

tagPROPVARIANT._fields_ = [
    ('_Union_1', _Union_1),
]


_VARIANT_NAME_4 = 'brecVal'

class __tagBRECORD(ctypes.Structure):
    _fields_ = [
        ('pvRecord', PVOID),
        ('pRecInfo', POINTER(IRecordInfo))
    ]

tagBRECORD = __tagBRECORD

# =======================  VARIANT  =======================

class _Union_2(ctypes.Union):
    class _tagVARIANT(ctypes.Structure):
        class _Union_3(ctypes.Union):
            _fields_ = [
                ('llVal', LONGLONG),  # VT_I8
                ('lVal', LONG),  # VT_I4
                ('bVal', BYTE),  # VT_UI1
                ('iVal', SHORT),  # VT_I2
                ('fltVal', FLOAT),  # VT_R4
                ('dblVal', DOUBLE),  # VT_R8
                ('boolVal', VARIANT_BOOL),  # VT_BOOL
                ('__OBSOLETE__VARIANT_BOOL', VARIANT_BOOL),  #
                ('scode', SCODE),  # VT_ERROR
                ('cyVal', CY),  # VT_CY
                ('date', DOUBLE),  # VT_DATE
                ('bstrVal', BSTR),  # VT_BSTR
                ('punkVal', POINTER(IUnknown)),  # VT_UNKNOWN
                ('pdispVal', POINTER(IDispatch)),  # VT_DISPATCH
                ('parray', LPSAFEARRAY),  # VT_ARRAY | VT_*
                ('pbVal', POINTER(UBYTE)),  # VT_BYREF | VT_UI1
                ('piVal', POINTER(SHORT)),  # VT_BYREF | VT_I2
                ('plVal', POINTER(LONG)),  # VT_BYREF | VT_I4
                ('pllVal', POINTER(LONGLONG)),  # VT_BYREF | VT_I4
                ('pfltVal', POINTER(FLOAT)),  # VT_BYREF | VT_R4
                ('pdblVal', POINTER(DOUBLE)),  # VT_BYREF | VT_R8
                ('pboolVal', POINTER(VARIANT_BOOL)),  # VT_BYREF | VT_BOOL
                ('__OBSOLETE__VARIANT_PBOOL', POINTER(VARIANT_BOOL)),  # VT_BYREF | VT_BOOL
                ('pscode', POINTER(SCODE)),  # VT_BYREF | VT_ERROR
                ('pcyVal', POINTER(CY)),  # VT_BYREF | VT_CY
                ('pdate', POINTER(DOUBLE)),  # VT_BYREF | VT_DATE
                ('pbstrVal', POINTER(BSTR)),  # VT_BYREF | VT_BSTR
                ('ppunkVal', POINTER(POINTER(IUnknown))),  # VT_BYREF | VT_UNKNOWN
                ('ppdispVal', POINTER(POINTER(IDispatch))),  # VT_BYREF | VT_DISPATCH
                ('pparray', POINTER(LPSAFEARRAY)),  # VT_BYREF | VT_ARRAY | VT_*
                ('pvarVal', POINTER(VARIANT)),  # VT_BYREF | VT_VARIANT
                ('byref', PVOID),
                ('cVal', CHAR),  # VT_I1
                ('uiVal', USHORT),  # VT_UI2
                ('ulVal', ULONG),  # VT_UI4
                ('ullVal', ULONGLONG),  # VT_UI8
                ('intVal', INT),  # VT_INT
                ('uintVal', UINT),  # VT_UINT
                ('pdecVal', POINTER(DECIMAL)),  # VT_BYREF | VT_DECIMAL
                ('pcVal', POINTER(CHAR)),  # VT_BYREF | VT_I1
                ('puiVal', POINTER(USHORT)),  # VT_BYREF | VT_UI2
                ('pulVal', POINTER(ULONG)),  # VT_BYREF | VT_UI4
                ('pullVal', POINTER(ULONGLONG)),  # VT_BYREF | VT_UI8
                ('pintVal', POINTER(INT)),  # VT_BYREF | VT_INT
                ('puintVal', POINTER(UINT)),  # VT_BYREF | VT_UINT
                (_VARIANT_NAME_4, tagBRECORD),
            ]

        _fields_ = [
            ('vt', VARTYPE),
            ('wReserved1', PROPVAR_PAD1),
            ('wReserved2', PROPVAR_PAD2),
            ('wReserved3', PROPVAR_PAD3),
            ('Union_3', _Union_3),
        ]

        _anonymous_ = ('Union_3',)

    _fields_ = [
        ('tagVARIANT', _tagVARIANT),
        ('decVal', DECIMAL),
    ]

    _anonymous_ = ('tagVARIANT',)




tagVARIANT._fields_ = [
    ('Union_2', _Union_2),
]
tagVARIANT._anonymous_ = ('Union_2',)




# =============================================================

# ===================  Interface Methods  =====================

IEnumSTATSTG._methods_ = [
    COMMETHOD(
        [helpstring('Method Next'), 'local'],
        HRESULT,
        'Next',
        (['in'], ULONG, 'celt'),
        (['out'], POINTER(STATSTG), 'rgelt'),
        (['out'], POINTER(ULONG), 'pceltFetched'),
    ),
    COMMETHOD(
        [helpstring('Method Skip')],
        HRESULT,
        'Skip',
        (['in'], ULONG, 'celt'),
    ),
    COMMETHOD(
        [helpstring('Method Reset')],
        HRESULT,
        'Reset',
    ),
    COMMETHOD(
        [helpstring('Method Clone')],
        HRESULT,
        'Clone',
        (['out'], POINTER(POINTER(IEnumSTATSTG)), 'ppenum'),
    ),
]

IStorage._methods_ = [
    COMMETHOD(
        [helpstring('Method CreateStream')],
        HRESULT,
        'CreateStream',
        (['in'], POINTER(OLECHAR), 'pwcsName'),
        (['in'], DWORD, 'grfMode'),
        (['in'], DWORD, 'reserved1'),
        (['in'], DWORD, 'reserved2'),
        (['out'], POINTER(POINTER(IStream)), 'ppstm'),
    ),
    COMMETHOD(
        [helpstring('Method OpenStream'), 'local'],
        HRESULT,
        'OpenStream',
        (['in'], POINTER(OLECHAR), 'pwcsName'),
        (['in'], POINTER(VOID), 'reserved1'),
        (['in'], DWORD, 'grfMode'),
        (['in'], DWORD, 'reserved2'),
        (['out'], POINTER(POINTER(IStream)), 'ppstm'),
    ),
    COMMETHOD(
        [helpstring('Method CreateStorage')],
        HRESULT,
        'CreateStorage',
        (['in'], POINTER(OLECHAR), 'pwcsName'),
        (['in'], DWORD, 'grfMode'),
        (['in'], DWORD, 'reserved1'),
        (['in'], DWORD, 'reserved2'),
        (['out'], POINTER(POINTER(IStorage)), 'ppstg'),
    ),
    COMMETHOD(
        [helpstring('Method OpenStorage')],
        HRESULT,
        'OpenStorage',
        (['unique', 'in'], POINTER(OLECHAR), 'pwcsName'),
        (['unique', 'in'], POINTER(IStorage), 'pstgPriority'),
        (['in'], DWORD, 'grfMode'),
        (['unique', 'in'], SNB, 'snbExclude'),
        (['in'], DWORD, 'reserved'),
        (['out'], POINTER(POINTER(IStorage)), 'ppstg'),
    ),
    COMMETHOD(
        [helpstring('Method CopyTo'), 'local'],
        HRESULT,
        'CopyTo',
        (['in'], DWORD, 'ciidExclude'),
        (['in'], POINTER(IID), 'rgiidExclude'),
        (['in'], SNB, 'snbExclude'),
        (['in'], POINTER(IStorage), 'pstgDest'),
    ),
    COMMETHOD(
        [helpstring('Method MoveElementTo')],
        HRESULT,
        'MoveElementTo',
        (['in'], POINTER(OLECHAR), 'pwcsName'),
        (['unique', 'in'], POINTER(IStorage), 'pstgDest'),
        (['in'], POINTER(OLECHAR), 'pwcsNewName'),
        (['in'], DWORD, 'grfFlags'),
    ),
    COMMETHOD(
        [helpstring('Method Commit')],
        HRESULT,
        'Commit',
        (['in'], DWORD, 'grfCommitFlags'),
    ),
    COMMETHOD(
        [helpstring('Method Revert')],
        HRESULT,
        'Revert',
    ),
    COMMETHOD(
        [helpstring('Method EnumElements'), 'local'],
        HRESULT,
        'EnumElements',
        (['in'], DWORD, 'reserved1'),
        (['in'], POINTER(VOID), 'reserved2'),
        (['in'], DWORD, 'reserved3'),
        (['out'], POINTER(POINTER(IEnumSTATSTG)), 'ppenum'),
    ),
    COMMETHOD(
        [helpstring('Method DestroyElement')],
        HRESULT,
        'DestroyElement',
        (['in'], POINTER(OLECHAR), 'pwcsName'),
    ),
    COMMETHOD(
        [helpstring('Method RenameElement')],
        HRESULT,
        'RenameElement',
        (['in'], POINTER(OLECHAR), 'pwcsOldName'),
        (['in'], POINTER(OLECHAR), 'pwcsNewName'),
    ),
    COMMETHOD(
        [helpstring('Method SetElementTimes')],
        HRESULT,
        'SetElementTimes',
        (['unique', 'in'], POINTER(OLECHAR), 'pwcsName'),
        (['unique', 'in'], POINTER(FILETIME), 'pctime'),
        (['unique', 'in'], POINTER(FILETIME), 'patime'),
        (['unique', 'in'], POINTER(FILETIME), 'pmtime'),
    ),
    COMMETHOD(
        [helpstring('Method SetClass')],
        HRESULT,
        'SetClass',
        (['in'], REFCLSID, 'clsid'),
    ),
    COMMETHOD(
        [helpstring('Method SetStateBits')],
        HRESULT,
        'SetStateBits',
        (['in'], DWORD, 'grfStateBits'),
        (['in'], DWORD, 'grfMask'),
    ),
    COMMETHOD(
        [helpstring('Method Stat')],
        HRESULT,
        'Stat',
        (['out'], POINTER(STATSTG), 'pstatstg'),
        (['in'], DWORD, 'grfStatFlag'),
    ),
]

ISequentialStream._methods_ = [
    COMMETHOD(
        [helpstring('Method Read'), 'local'],
        HRESULT,
        'Read',
        (['out'], POINTER(VOID), 'pv'),
        (['in'], ULONG, 'cb'),
        (['out'], POINTER(ULONG), 'pcbRead'),
    ),
    COMMETHOD(
        [helpstring('Method Write'), 'local'],
        HRESULT,
        'Write',
        (['in'], POINTER(VOID), 'pv'),
        (['in'], ULONG, 'cb'),
        (['out'], POINTER(ULONG), 'pcbWritten'),
    ),
]

IStream._methods_ = [
    COMMETHOD(
        [helpstring('Method Seek'), 'local'],
        HRESULT,
        'Seek',
        (['in'], LARGE_INTEGER, 'dlibMove'),
        (['in'], DWORD, 'dwOrigin'),
        (
            ['out'],
            POINTER(ULARGE_INTEGER),
            'plibNewPosition'
        ),
    ),
    COMMETHOD(
        [helpstring('Method SetSize')],
        HRESULT,
        'SetSize',
        (['in'], ULARGE_INTEGER, 'libNewSize'),
    ),
    COMMETHOD(
        [helpstring('Method CopyTo'), 'local'],
        HRESULT,
        'CopyTo',
        (['in'], POINTER(IStream), 'pstm'),
        (['in'], ULARGE_INTEGER, 'cb'),
        (['out'], POINTER(ULARGE_INTEGER), 'pcbRead'),
        (['out'], POINTER(ULARGE_INTEGER), 'pcbWritten'),
    ),
    COMMETHOD(
        [helpstring('Method Commit')],
        HRESULT,
        'Commit',
        (['in'], DWORD, 'grfCommitFlags'),
    ),
    COMMETHOD(
        [helpstring('Method Revert')],
        HRESULT,
        'Revert',
    ),
    COMMETHOD(
        [helpstring('Method LockRegion')],
        HRESULT,
        'LockRegion',
        (['in'], ULARGE_INTEGER, 'libOffset'),
        (['in'], ULARGE_INTEGER, 'cb'),
        (['in'], DWORD, 'dwLockType'),
    ),
    COMMETHOD(
        [helpstring('Method UnlockRegion')],
        HRESULT,
        'UnlockRegion',
        (['in'], ULARGE_INTEGER, 'libOffset'),
        (['in'], ULARGE_INTEGER, 'cb'),
        (['in'], DWORD, 'dwLockType'),
    ),
    COMMETHOD(
        [helpstring('Method Stat')],
        HRESULT,
        'Stat',
        (['out'], POINTER(STATSTG), 'pstatstg'),
        (['in'], DWORD, 'grfStatFlag'),
    ),
    COMMETHOD(
        [helpstring('Method Clone')],
        HRESULT,
        'Clone',
        (['out'], POINTER(POINTER(IStream)), 'ppstm'),
    ),
]

IEnumSTATPROPSTG._methods_ = [
    COMMETHOD(
        [helpstring('Method Next'), 'local'],
        HRESULT,
        'Next',
        (['in'], ULONG, 'celt'),
        (['out'], POINTER(STATPROPSTG), 'rgelt'),
        (['out'], POINTER(ULONG), 'pceltFetched'),
    ),
    COMMETHOD(
        [helpstring('Method Skip')],
        HRESULT,
        'Skip',
        (['in'], ULONG, 'celt'),
    ),
    COMMETHOD(
        [helpstring('Method Reset')],
        HRESULT,
        'Reset',
    ),
    COMMETHOD(
        [helpstring('Method Clone')],
        HRESULT,
        'Clone',
        (
            ['out'],
            POINTER(POINTER(IEnumSTATPROPSTG)),
            'ppenum'
        ),
    ),
]

IEnumSTATPROPSETSTG._methods_ = [
    COMMETHOD(
        [helpstring('Method Next'), 'local'],
        HRESULT,
        'Next',
        (['in'], ULONG, 'celt'),
        (['out'], POINTER(STATPROPSETSTG), 'rgelt'),
        (['out'], POINTER(ULONG), 'pceltFetched'),
    ),
    COMMETHOD(
        [helpstring('Method Skip')],
        HRESULT,
        'Skip',
        (['in'], ULONG, 'celt'),
    ),
    COMMETHOD(
        [helpstring('Method Reset')],
        HRESULT,
        'Reset',
    ),
    COMMETHOD(
        [helpstring('Method Clone')],
        HRESULT,
        'Clone',
        (
            ['out'],
            POINTER(POINTER(IEnumSTATPROPSETSTG)),
            'ppenum'
        ),
    ),
]

# =============================================================

# Property IDs for the DiscardableInformation Property Set
PIDDI_THUMBNAIL = 0x00000002  # VT_BLOB

# Property IDs for the SummaryInformation Property Set
PIDSI_TITLE = 0x00000002  # VT_LPSTR
PIDSI_SUBJECT = 0x00000003  # VT_LPSTR
PIDSI_AUTHOR = 0x00000004  # VT_LPSTR
PIDSI_KEYWORDS = 0x00000005  # VT_LPSTR
PIDSI_COMMENTS = 0x00000006  # VT_LPSTR
PIDSI_TEMPLATE = 0x00000007  # VT_LPSTR
PIDSI_LASTAUTHOR = 0x00000008  # VT_LPSTR
PIDSI_REVNUMBER = 0x00000009  # VT_LPSTR
PIDSI_EDITTIME = 0x0000000A  # VT_FILETIME (UTC)
PIDSI_LASTPRINTED = 0x0000000B  # VT_FILETIME (UTC)
PIDSI_CREATE_DTM = 0x0000000C  # VT_FILETIME (UTC)
PIDSI_LASTSAVE_DTM = 0x0000000D  # VT_FILETIME (UTC)
PIDSI_PAGECOUNT = 0x0000000E  # VT_I4
PIDSI_WORDCOUNT = 0x0000000F  # VT_I4
PIDSI_CHARCOUNT = 0x00000010  # VT_I4
PIDSI_THUMBNAIL = 0x00000011  # VT_CF
PIDSI_APPNAME = 0x00000012  # VT_LPSTR
PIDSI_DOC_SECURITY = 0x00000013  # VT_I4

# Property IDs for the DocSummaryInformation Property Set
PIDDSI_CATEGORY = 0x00000002  # VT_LPSTR
PIDDSI_PRESFORMAT = 0x00000003  # VT_LPSTR
PIDDSI_BYTECOUNT = 0x00000004  # VT_I4
PIDDSI_LINECOUNT = 0x00000005  # VT_I4
PIDDSI_PARCOUNT = 0x00000006  # VT_I4
PIDDSI_SLIDECOUNT = 0x00000007  # VT_I4
PIDDSI_NOTECOUNT = 0x00000008  # VT_I4
PIDDSI_HIDDENCOUNT = 0x00000009  # VT_I4
PIDDSI_MMCLIPCOUNT = 0x0000000A  # VT_I4
PIDDSI_SCALE = 0x0000000B  # VT_BOOL
PIDDSI_HEADINGPAIR = 0x0000000C  # VT_VARIANT | VT_VECTOR
PIDDSI_DOCPARTS = 0x0000000D  # VT_LPSTR | VT_VECTOR
PIDDSI_MANAGER = 0x0000000E  # VT_LPSTR
PIDDSI_COMPANY = 0x0000000F  # VT_LPSTR
PIDDSI_LINKSDIRTY = 0x00000010  # VT_BOOL

# FMTID_MediaFileSummaryInfo - Property IDs
PIDMSI_EDITOR = 0x00000002  # VT_LPWSTR
PIDMSI_SUPPLIER = 0x00000003  # VT_LPWSTR
PIDMSI_SOURCE = 0x00000004  # VT_LPWSTR
PIDMSI_SEQUENCE_NO = 0x00000005  # VT_LPWSTR
PIDMSI_PROJECT = 0x00000006  # VT_LPWSTR
PIDMSI_STATUS = 0x00000007  # VT_UI4
PIDMSI_OWNER = 0x00000008  # VT_LPWSTR
PIDMSI_RATING = 0x00000009  # VT_LPWSTR
PIDMSI_PRODUCTION = 0x0000000A  # VT_FILETIME (UTC)
PIDMSI_COPYRIGHT = 0x0000000B  # VT_LPWSTR

# =======================  FUNCTIONS  =========================


ole32 = ctypes.windll.OLE32

# _Check_return_ WINOLEAPI PropVariantCopy(
#     _Out_ PROPVARIANT* pvarDest,
#     _In_ PROPVARIANT * pvarSrc
# );
PropVariantCopy = ole32.PropVariantCopy
PropVariantCopy.restype = HRESULT

# WINOLEAPI PropVariantClear(_Inout_ PROPVARIANT* pvar);
PropVariantClear = ole32.PropVariantClear
PropVariantClear.restype = HRESULT

# WINOLEAPI FreePropVariantArray(
#     _In_ c_ulong cVariants,
#     _Inout_updates_(cVariants) PROPVARIANT* rgvars
# );
FreePropVariantArray = ole32.FreePropVariantArray
FreePropVariantArray.restype = HRESULT


def PropVariantInit(pvar):
    return ctypes.memset(pvar, 0, ctypes.sizeof(PROPVARIANT))


# EXTERN_C
# _Success_(TRUE) // Raises status on failure
# SERIALIZEDPROPERTYVALUE* __stdcall
# StgConvertVariantToProperty(
#     _In_ PROPVARIANT* pvar,
#     _In_ c_ushort CodePage,
#     _Out_writes_bytes_opt_(*pcb) SERIALIZEDPROPERTYVALUE* pprop,
#     _Inout_ c_ulong* pcb,
#     _In_ PROPID pid,
#     _Reserved_ BOOLEAN fReserved,
#     _Inout_opt_ c_ulong* pcIndirect
# );
StgConvertVariantToProperty = ole32.StgConvertVariantToProperty
StgConvertVariantToProperty.restype = SERIALIZEDPROPERTYVALUE

# EXTERN_C
# _Success_(TRUE) // Raises status on failure
# BOOLEAN __stdcall
# StgConvertPropertyToVariant(
#     _In_ SERIALIZEDPROPERTYVALUE* pprop,
#     _In_ c_ushort CodePage,
#     _Out_ PROPVARIANT* pvar,
#     _In_ PMemoryAllocator* pma
# );
StgConvertPropertyToVariant = ole32.StgConvertPropertyToVariant
StgConvertPropertyToVariant.restype = BOOLEAN

# Additional Prototypes for ALL interfaces
oleaut32 = ctypes.windll.OLEAUT32

# c_ulong BSTR_UserSize(
#     __RPC__in c_ulong *,
#     c_ulong,
#     __RPC__in BSTR *
# );
BSTR_UserSize = oleaut32.BSTR_UserSize
BSTR_UserSize.restype = ULONG

# c_ubyte * BSTR_UserMarshal(
#     __RPC__in c_ulong *,
#     __RPC__inout_xcount(0) c_ubyte *,
#     __RPC__in BSTR *
# );
BSTR_UserMarshal = oleaut32.BSTR_UserMarshal
BSTR_UserMarshal.restype = POINTER(UBYTE)

# c_ubyte * BSTR_UserUnmarshal(
#     __RPC__in c_ulong *,
#     __RPC__in_xcount(0) c_ubyte *,
#     __RPC__out BSTR *
# );
BSTR_UserUnmarshal = oleaut32.BSTR_UserUnmarshal
BSTR_UserUnmarshal.restype = POINTER(UBYTE)

# VOID BSTR_UserFree(
#     __RPC__in c_ulong *,
#     __RPC__in BSTR *
# );
BSTR_UserFree = oleaut32.BSTR_UserFree
BSTR_UserFree.restype = VOID

# c_ulong LPSAFEARRAY_UserSize(
#     __RPC__in c_ulong *,
#     c_ulong,
#     __RPC__in LPSAFEARRAY *
# );
LPSAFEARRAY_UserSize = oleaut32.LPSAFEARRAY_UserSize
LPSAFEARRAY_UserSize.restype = ULONG

# c_ubyte * LPSAFEARRAY_UserMarshal(
#     __RPC__in c_ulong *,
#     __RPC__inout_xcount(0) c_ubyte *,
#     __RPC__in LPSAFEARRAY *
# );
LPSAFEARRAY_UserMarshal = oleaut32.LPSAFEARRAY_UserMarshal
LPSAFEARRAY_UserMarshal.restype = POINTER(UBYTE)

# c_ubyte * LPSAFEARRAY_UserUnmarshal(
#     __RPC__in c_ulong *,
#     __RPC__in_xcount(0) c_ubyte *,
#     __RPC__out LPSAFEARRAY *
# );
LPSAFEARRAY_UserUnmarshal = oleaut32.LPSAFEARRAY_UserUnmarshal
LPSAFEARRAY_UserUnmarshal.restype = POINTER(UBYTE)

# VOID LPSAFEARRAY_UserFree(
#     __RPC__in c_ulong *,
#     __RPC__in LPSAFEARRAY *
# );
LPSAFEARRAY_UserFree = oleaut32.LPSAFEARRAY_UserFree
LPSAFEARRAY_UserFree.restype = VOID

# c_ulong BSTR_UserSize64(
#     __RPC__in c_ulong *,
#     c_ulong,
#     __RPC__in BSTR *
# );
BSTR_UserSize64 = oleaut32.BSTR_UserSize64
BSTR_UserSize64.restype = ULONG

# c_ubyte * BSTR_UserMarshal64(
#     __RPC__in c_ulong *,
#     __RPC__inout_xcount(0) c_ubyte *,
#     __RPC__in BSTR *
# );
BSTR_UserMarshal64 = oleaut32.BSTR_UserMarshal64
BSTR_UserMarshal64.restype = POINTER(UBYTE)

# c_ubyte * BSTR_UserUnmarshal64(
#     __RPC__in c_ulong *,
#     __RPC__in_xcount(0) c_ubyte *,
#     __RPC__out BSTR *
# );
BSTR_UserUnmarshal64 = oleaut32.BSTR_UserUnmarshal64
BSTR_UserUnmarshal64.restype = POINTER(UBYTE)

# VOID BSTR_UserFree64(
#     __RPC__in c_ulong *,
#     __RPC__in BSTR *
# );
BSTR_UserFree64 = oleaut32.BSTR_UserFree64
BSTR_UserFree64.restype = VOID

# c_ulong LPSAFEARRAY_UserSize64(
#     __RPC__in c_ulong *,
#     c_ulong,
#     __RPC__in LPSAFEARRAY *
# );
LPSAFEARRAY_UserSize64 = oleaut32.LPSAFEARRAY_UserSize64
LPSAFEARRAY_UserSize64.restype = ULONG

# c_ubyte * LPSAFEARRAY_UserMarshal64(
#     __RPC__in c_ulong *,
#     __RPC__inout_xcount(0) c_ubyte *,
#     __RPC__in LPSAFEARRAY *
# );
LPSAFEARRAY_UserMarshal64 = oleaut32.LPSAFEARRAY_UserMarshal64
LPSAFEARRAY_UserMarshal64.restype = POINTER(UBYTE)

# c_ubyte * LPSAFEARRAY_UserUnmarshal64(
#     __RPC__in c_ulong *,
#     __RPC__in_xcount(0) c_ubyte *,
#     __RPC__out LPSAFEARRAY *
# );
LPSAFEARRAY_UserUnmarshal64 = oleaut32.LPSAFEARRAY_UserUnmarshal64
LPSAFEARRAY_UserUnmarshal64.restype = POINTER(UBYTE)

# VOID LPSAFEARRAY_UserFree64(
#     __RPC__in c_ulong *,
#     __RPC__in LPSAFEARRAY *
# );
LPSAFEARRAY_UserFree64 = oleaut32.LPSAFEARRAY_UserFree64
LPSAFEARRAY_UserFree64.restype = VOID

# HRESULT PropVariantClear(
#   [in] PROPVARIANT *pvar
# );
_PropVariantClear = ole32.PropVariantClear
_PropVariantClear.restype = HRESULT

# HRESULT PropVariantCopy(
#   [in, out] PROPVARIANT       *pvarDest,
#   [in]      const PROPVARIANT *pvarSrc
# );
_PropVariantCopy = ole32.PropVariantCopy
_PropVariantCopy.restype = HRESULT

propsys = ctypes.windll.Propsys

# PSSTDAPI PropVariantChangeType(
#   [out] PROPVARIANT          *ppropvarDest,
#   [in]  REFPROPVARIANT       propvarSrc,
#   [in]  PROPVAR_CHANGE_FLAGS flags,
#   [in]  VARTYPE              vt
# );
_PropVariantChangeType = propsys.PropVariantChangeType
_PropVariantChangeType.restype = HRESULT

# =============================================================
