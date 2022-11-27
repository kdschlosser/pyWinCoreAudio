from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib

CLSID_DPCEE3RApi = CLSID('{248F37DE-5F40-4DC6-9B51-3A14C1D71FD5}')
IID_IDPCEE3RApi = IID('{3DBC6EC3-E7BE-4168-A2A7-684F69ADD8DC}')


class IDPCEE3RApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IDPCEE3RApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetAE')],
            HRESULT,
            'GetAE',
            (['in'], LPWSTR, 'pwstrId'),
            (['out', 'retval'], POINTER(INT), 'pfAE')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSurr')],
            HRESULT,
            'GetSurr',
            (['in'], LPWSTR, 'pwstrId'),
            (['out', 'retval'], POINTER(INT), 'pfSurr')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetAE')],
            HRESULT,
            'SetAE',
            (['in'], LPWSTR, 'pwstrId'),
            (['in'], INT, 'fAE')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetSurr')],
            HRESULT,
            'SetSurr',
            (['in'], LPWSTR, 'pwstrId'),
            (['in'], INT, 'fSurr')
        ),
    ]


class DPCEE3RApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_DPCEE3RApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IDPCEE3RApi]

    @staticmethod
    def IDPCEE3RApi() -> IDPCEE3RApi:
        return comtypes.CoCreateInstance(
            CLSID_DPCEE3RApi,
            IDPCEE3RApi,
            comtypes.CLSCTX_ALL
        )

