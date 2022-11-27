from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_DPCEE3Api = CLSID('{3A2B839D-6060-4711-B3A8-35132A881E78}')
IID_IDPCEE3Api = IID('{E3220AD1-CE0A-4036-8AA9-B6A07BD7A129}')


class IDPCEE3Api(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IDPCEE3Api
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method SupportPCEE3')],
            HRESULT,
            'SupportPCEE3',
            (['out', 'retval'], POINTER(INT), 'pfSupport')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SupportPCEE3EA')],
            HRESULT,
            'SupportPCEE3EA',
            (['out', 'retval'], POINTER(INT), 'pfSupport')
        ),
    ]


class DPCEE3Api(comtypes.CoClass):
    _reg_clsid_ = CLSID_DPCEE3Api
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IDPCEE3Api]

    @staticmethod
    def IDPCEE3Api() -> IDPCEE3Api:
        return comtypes.CoCreateInstance(
            CLSID_DPCEE3Api,
            IDPCEE3Api,
            comtypes.CLSCTX_ALL
        )