from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib

CLSID_RtkRecNoLimitApi = CLSID('{88A0B959-0AAC-486E-90EC-90444C40F7EC}')
IID_IRtkRecNoLimitApi = IID('{C69D6BA6-7141-475B-B301-518B43BE740D}')


class IRtkRecNoLimitApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IRtkRecNoLimitApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method SetRtkRecNoLimit')],
            HRESULT,
            'SetRtkRecNoLimit',
            (['in'], INT, 'fNoLimit')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetRtkRecNoLimit')],
            HRESULT,
            'GetRtkRecNoLimit',
            (['out', 'retval'], POINTER(INT), 'pfNoLimit')
        ),
    ]


class RtkRecNoLimitApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_RtkRecNoLimitApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IRtkRecNoLimitApi]

    @staticmethod
    def IRtkRecNoLimitApi() -> IRtkRecNoLimitApi:
        return comtypes.CoCreateInstance(
            CLSID_RtkRecNoLimitApi,
            IRtkRecNoLimitApi,
            comtypes.CLSCTX_ALL
        )
