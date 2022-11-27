from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_MCEQApi = CLSID('{498A0C2A-D4C5-4CBB-B62F-FB075A02F5A8}')
IID_IMCEQApi = IID('{D486701C-55CE-47C8-BBBA-5218D210502A}')


class IMCEQApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IMCEQApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method SetCurrentCEQ')],
            HRESULT,
            'SetCurrentCEQ',
            (['in'], INT, 'nCeqIndex')
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


class MCEQApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_MCEQApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IMCEQApi]

    @staticmethod
    def IMCEQApi() -> IMCEQApi:
        return comtypes.CoCreateInstance(
            CLSID_MCEQApi,
            IMCEQApi,
            comtypes.CLSCTX_ALL
        )

