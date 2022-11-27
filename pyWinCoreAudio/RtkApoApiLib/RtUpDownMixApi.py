from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_RtUpDownMixApi = CLSID('{A5A63314-41AC-484F-85B8-FC15335C24CF}')
IID_IRtUpDownMixApi = IID('{90361AEF-24CF-43B4-A5F2-EA00BE4533BA}')


class IRtUpDownMixApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IRtUpDownMixApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetUpDownMix')],
            HRESULT,
            'GetUpDownMix',
            (['in'], LPWSTR, 'pwszDeviceId'),
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetUpDownMix')],
            HRESULT,
            'SetUpDownMix',
            (['in'], LPWSTR, 'pwszDeviceId'),
            (['in'], LONG, 'fEnable')
        ),
    ]


class RtUpDownMixApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_RtUpDownMixApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IRtUpDownMixApi]

    @staticmethod
    def IRtUpDownMixApi() -> IRtUpDownMixApi:
        return comtypes.CoCreateInstance(
            CLSID_RtUpDownMixApi,
            IRtUpDownMixApi,
            comtypes.CLSCTX_ALL
        )
