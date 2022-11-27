
from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_MDRCApi = CLSID('{E51EF0ED-35D5-4BDE-93E9-28CA62B1E704}')
IID_IMDRCApi = IID('{4946741F-9170-468D-939E-0DABCF92D2D2}')


class IMDRCApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IMDRCApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method SetCurrentDRC')],
            HRESULT,
            'SetCurrentDRC',
            (['in'], INT, 'nDrcIndex')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetEpCnt')],
            HRESULT,
            'GetEpCnt',
            (['out', 'retval'], POINTER(INT), 'pnEpCnt')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetEpStatus')],
            HRESULT,
            'GetEpStatus',
            (['in'], INT, 'nEpIndex'),
            (['out', 'retval'], POINTER(WSTRING), 'ppstrEpName')
        ),
    ]


class MDRCApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_MDRCApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IMDRCApi]

    @staticmethod
    def IMDRCApi() -> IMDRCApi:
        return comtypes.CoCreateInstance(
            CLSID_MDRCApi,
            IMDRCApi,
            comtypes.CLSCTX_ALL
        )
