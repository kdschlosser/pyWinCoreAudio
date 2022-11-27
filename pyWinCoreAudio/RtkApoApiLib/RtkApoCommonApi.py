from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_RtkApoCommonApi = CLSID('{68E6DF30-3484-489B-8120-3C4C0256A758}')
IID_IRtkApoCommonApi = IID('{3F262F1A-8B1F-4082-BA80-A8C793C5E561}')


class IRtkApoCommonApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IRtkApoCommonApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method Set_DisableSysfx')],
            HRESULT,
            'Set_DisableSysfx',
            (['in'], LPWSTR, 'pwstrId'),
            (['in'], INT, 'fDisable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method Get_DisableSysfx')],
            HRESULT,
            'Get_DisableSysfx',
            (['in'], LPWSTR, 'pwstrId'),
            (['out', 'retval'], POINTER(INT), 'pfDisable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method Set_DisableSpkrSysfx')],
            HRESULT,
            'Set_DisableSpkrSysfx',
            (['in'], INT, 'fDisable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method Get_DisableSpkrSysfx')],
            HRESULT,
            'Get_DisableSpkrSysfx',
            (['out', 'retval'], POINTER(INT), 'pfDisable')
        ),
    ]


class RtkApoCommonApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_RtkApoCommonApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IRtkApoCommonApi]

    @staticmethod
    def IRtkApoCommonApi() -> IRtkApoCommonApi:
        return comtypes.CoCreateInstance(
            CLSID_RtkApoCommonApi,
            IRtkApoCommonApi,
            comtypes.CLSCTX_ALL
        )
