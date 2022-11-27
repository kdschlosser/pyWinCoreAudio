from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_RtkRecOptionsApi = CLSID('{31929B1C-6E3C-431C-A20D-2A2A8315344D}')
IID_IRtkRecOptionsApi = IID('{99E653C9-7EA5-4BDF-A759-48DEF7917459}')


class IRtkRecOptionsApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IRtkRecOptionsApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method SetNoLimit')],
            HRESULT,
            'SetNoLimit',
            (['in'], INT, 'fNoLimit')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetNoLimit')],
            HRESULT,
            'GetNoLimit',
            (['out', 'retval'], POINTER(INT), 'pfNoLimit')
        ),
    ]


class RtkRecOptionsApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_RtkRecOptionsApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IRtkRecOptionsApi]

    @staticmethod
    def IRtkRecOptionsApi() -> IRtkRecOptionsApi:
        return comtypes.CoCreateInstance(
            CLSID_RtkRecOptionsApi,
            IRtkRecOptionsApi,
            comtypes.CLSCTX_ALL
        )
