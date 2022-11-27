
from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib

CLSID_DolbyApi = CLSID('{6201C0B0-A1DC-4AE2-9665-8DAB2E380073}')
IID_IDolbyApi = IID('{4EB57548-5257-4765-8C29-17C4E6A3FF00}')


class IDolbyApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IDolbyApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method Initialize')],
            HRESULT,
            'Initialize'
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetDolbyPCEE')],
            HRESULT,
            'GetDolbyPCEE',
            (['in'], LPWSTR, 'pwszDeviceId'),
            (['out'], POINTER(ULONG), 'dwDolbyPCEE')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetHeadphoneMode')],
            HRESULT,
            'GetHeadphoneMode',
            (['in'], LPWSTR, 'pwszDeviceId'),
            (['out'], POINTER(LONG), 'pfHeadphoneMode')
        ),
    ]


class DolbyApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_DolbyApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IDolbyApi]

    @staticmethod
    def IDolbyApi() -> IDolbyApi:
        return comtypes.CoCreateInstance(
            CLSID_DolbyApi,
            IDolbyApi,
            comtypes.CLSCTX_ALL
        )
