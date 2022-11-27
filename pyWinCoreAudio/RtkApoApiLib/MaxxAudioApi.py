from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_MaxxAudioApi = GUID('{5B05A596-33CA-4D65-B6C8-42704CA9BE2A}')
IID_IMaxxAudioApi = IID('{B8C3502C-5BAC-4B01-AB47-E3E9BFEA7D0C}')


class IMaxxAudioApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IMaxxAudioApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']

    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetCurrentMode')],
            HRESULT,
            'GetCurrentMode',
            (['out', 'retval'], POINTER(INT), 'pnMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetCurrentMode')],
            HRESULT,
            'SetCurrentMode',
            (['in'], INT, 'nMode')
        ),
    ]


class MaxxAudioApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_MaxxAudioApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IMaxxAudioApi]

    @staticmethod
    def IMaxxAudioApi() -> IMaxxAudioApi:
        return comtypes.CoCreateInstance(
            CLSID_MaxxAudioApi,
            IMaxxAudioApi,
            comtypes.CLSCTX_ALL
        )
