from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib

CLSID_RtkBassBackApi = IID('{7ECCA61E-6028-4AFA-B03C-B68004412CAF}')
IID_IRtkBassBackApi = IID('{52DE4E5E-A3B8-4EFD-B8B4-3FA365B76D49}')


class IRtkBassBackApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IRtkBassBackApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method Initialize')],
            HRESULT,
            'Initialize',
            (['in'], LPWSTR, 'pwszDeviceId')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetEnable')],
            HRESULT,
            'SetEnable',
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetEnable')],
            HRESULT,
            'GetEnable',
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetLevel')],
            HRESULT,
            'SetLevel',
            (['in'], INT, 'nLevel')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetLevel')],
            HRESULT,
            'GetLevel',
            (['out', 'retval'], POINTER(INT), 'pnLevel')
        ),
    ]


class RtkBassBackApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_RtkBassBackApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IRtkBassBackApi]

    @staticmethod
    def IRtkBassBackApi() -> IRtkBassBackApi:
        return comtypes.CoCreateInstance(
            CLSID_RtkBassBackApi,
            IRtkBassBackApi,
            comtypes.CLSCTX_ALL
        )
