from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_RtkOmniSoundApi = CLSID('{7505CB57-3A77-46A4-8B04-75897B0AA643}')
IID_IRtkOmniSoundApi = IID('{4D950B23-AAFC-4A1D-88F6-E945DA2FAFF8}')


class IRtkOmniSoundApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IRtkOmniSoundApi
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
            [comtypes.helpstring('method SetCenterLevel')],
            HRESULT,
            'SetCenterLevel',
            (['in'], INT, 'nLevel')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetCenterLevel')],
            HRESULT,
            'GetCenterLevel',
            (['out', 'retval'], POINTER(INT), 'pnLevel')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetSurroundLevel')],
            HRESULT,
            'SetSurroundLevel',
            (['in'], INT, 'nLevel')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSurroundLevel')],
            HRESULT,
            'GetSurroundLevel',
            (['out', 'retval'], POINTER(INT), 'pnLevel')
        ),
    ]


class RtkOmniSoundApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_RtkOmniSoundApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IRtkOmniSoundApi]

    @staticmethod
    def IRtkOmniSoundApi() -> IRtkOmniSoundApi:
        return comtypes.CoCreateInstance(
            CLSID_RtkOmniSoundApi,
            IRtkOmniSoundApi,
            comtypes.CLSCTX_ALL
        )
