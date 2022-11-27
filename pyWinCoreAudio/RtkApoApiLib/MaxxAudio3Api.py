from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib

CLSID_MaxxAudio3Api = CLSID('{9D1C81FD-A77A-4CC0-BD6E-A827C9E38ACA}')
IID_IMaxxAudio3Api = IID('{A1FFF951-1D4F-41DF-8311-4AE5F2CE8B9C}')


class IMaxxAudio3Api(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IMaxxAudio3Api
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetCurrentMode')],
            HRESULT,
            'GetEnable',
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetCurrentMode')],
            HRESULT,
            'SetEnable',
            (['in'], LONG, 'fEnable')
        ),
    ]


class MaxxAudio3Api(comtypes.CoClass):
    _reg_clsid_ = CLSID_MaxxAudio3Api
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IMaxxAudio3Api]

    @staticmethod
    def IMaxxAudio3Api() -> IMaxxAudio3Api:
        return comtypes.CoCreateInstance(
            CLSID_MaxxAudio3Api,
            IMaxxAudio3Api,
            comtypes.CLSCTX_ALL
        )

