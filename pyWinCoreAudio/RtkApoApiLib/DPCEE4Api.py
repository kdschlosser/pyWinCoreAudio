
from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_DPCEE4Api = CLSID('{853B15F2-A9DE-4592-9FA9-7A69F6AEA6A2}')
IID_IDPCEE4Api = IID('{C2A289BD-AC54-4D32-BDB0-C7C4A8B95F79}')


class IDPCEE4Api(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IDPCEE4Api
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = []


class DPCEE4Api(comtypes.CoClass):
    _reg_clsid_ = CLSID_DPCEE4Api
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IDPCEE4Api]

    @staticmethod
    def IDPCEE4Api() -> IDPCEE4Api:
        return comtypes.CoCreateInstance(
            CLSID_DPCEE4Api,
            IDPCEE4Api,
            comtypes.CLSCTX_ALL
        )
