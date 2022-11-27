from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_DDLiveApi = CLSID('{9F39C0AE-979D-4BFC-890D-4BDF6FD7BDE1}')
IID_IDDLiveApi = IID('{434BAB46-BCF6-4887-9AC0-8F6A59FF98B8}')


class IDDLiveApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IDDLiveApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method Initialize')],
            HRESULT,
            'Initialize',
            (['in'], LPWSTR, 'pwszDeviceId')
        ),
        COMMETHOD(
            [comtypes.helpstring('method EnableDolbyDigitalLive')],
            HRESULT,
            'EnableDolbyDigitalLive',
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetDolbyDigitalLive')],
            HRESULT,
            'GetDolbyDigitalLive',
            (['out'], POINTER(LONG), 'pfEnable')
        ),
    ]


class DDLiveApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_DDLiveApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IDDLiveApi]

    @staticmethod
    def IDDLiveApi() -> IDDLiveApi:
        return comtypes.CoCreateInstance(
            CLSID_DDLiveApi,
            IDDLiveApi,
            comtypes.CLSCTX_ALL
        )
