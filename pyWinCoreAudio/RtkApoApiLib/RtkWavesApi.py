
from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib

CLSID_RtkWavesApi = CLSID('{F2464A7E-5B09-4199-826F-7D6A16AF2E59}')
IID_IRtkWavesApi = IID('{1887CDE7-4D3E-4111-A406-25264493F449}')


class IRtkWavesApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IRtkWavesApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method MASetOuputModeSize')],
            HRESULT,
            'MASetOuputModeSize',
            (['in'], LPWSTR, 'pwstrId'),
            (['in'], INT, 'nType'),
            (['in'], INT, 'nSize')
        ),
        COMMETHOD(
            [comtypes.helpstring('method MAGetOuputModeSize')],
            HRESULT,
            'MAGetOuputModeSize',
            (['in'], LPWSTR, 'pwstrId'),
            (['in'], INT, 'nType'),
            (['in', 'out'], POINTER(INT), 'pnSize')
        ),
    ]


class RtkWavesApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_RtkWavesApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IRtkWavesApi]

    @staticmethod
    def IRtkWavesApi() -> IRtkWavesApi:
        return comtypes.CoCreateInstance(
            CLSID_RtkWavesApi,
            IRtkWavesApi,
            comtypes.CLSCTX_ALL
        )
