from ..data_types import *
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_RtkSRSPSApi = CLSID('{42F7121A-9917-4927-9E44-C9B305B0EDBB}')
IID_IRtkSRSPSApi = IID('{558E95F6-4D04-43F7-B6C7-B398586AA8BE}')


class IRtkSRSPSApi(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IRtkSRSPSApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.dispid(1)],
            HRESULT,
            'Initialize',
            (['in'], LPWSTR, 'pwszDeviceId'),
            (['in'], INT, 'nProgramId')
        ),
        COMMETHOD(
            [comtypes.dispid(2)],
            HRESULT,
            'LoadPreset',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType')
        ),
        COMMETHOD(
            [comtypes.dispid(3)],
            HRESULT,
            'RestoreDefault',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType')
        ),
    ]


class RtkSRSPSApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_RtkSRSPSApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IRtkSRSPSApi]

    @staticmethod
    def IRtkSRSPSApi() -> IRtkSRSPSApi:
        return comtypes.CoCreateInstance(
            CLSID_RtkSRSPSApi,
            IRtkSRSPSApi,
            comtypes.CLSCTX_ALL
        )
