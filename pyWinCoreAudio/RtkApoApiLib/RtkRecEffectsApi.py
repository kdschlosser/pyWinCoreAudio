from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib

CLSID_RtkRecEffectsApi = CLSID('{BC1F9ACF-DCA1-43D4-AFEE-D742DC11473C}')
IID_IRtkRecEffectsApi = IID('{04DD16AD-F885-4849-85B8-6124525E040E}')


class IRtkRecEffectsApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IRtkRecEffectsApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetEndpointCount')],
            HRESULT,
            'GetEndpointCount',
            (['out', 'retval'], POINTER(INT), 'pcnt')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetModePKEYInfo')],
            HRESULT,
            'GetModePKEYInfo',
            (['in'], INT, 'nEndpointIndex'),
            (['out'], POINTER(WSTRING), 'psRegPath')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetBeamForming')],
            HRESULT,
            'SetBeamForming',
            (['in'], INT, 'nEndpointIndex'),
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetBeamForming')],
            HRESULT,
            'GetBeamForming',
            (['in'], INT, 'nEndpointIndex'),
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetFarFieldPickup')],
            HRESULT,
            'SetFarFieldPickup',
            (['in'], INT, 'nEndpointIndex'),
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetFarFieldPickup')],
            HRESULT,
            'GetFarFieldPickup',
            (['in'], INT, 'nEndpointIndex'),
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetAEC')],
            HRESULT,
            'SetAEC',
            (['in'], INT, 'nEndpointIndex'),
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetAEC')],
            HRESULT,
            'GetAEC',
            (['in'], INT, 'nEndpointIndex'),
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetNS')],
            HRESULT,
            'SetNS',
            (['in'], INT, 'nEndpointIndex'),
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetNS')],
            HRESULT,
            'GetNS',
            (['in'], INT, 'nEndpointIndex'),
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
    ]


class RtkRecEffectsApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_RtkRecEffectsApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IRtkRecEffectsApi]

    @staticmethod
    def IRtkRecEffectsApi() -> IRtkRecEffectsApi:
        return comtypes.CoCreateInstance(
            CLSID_RtkRecEffectsApi,
            IRtkRecEffectsApi,
            comtypes.CLSCTX_ALL
        )
