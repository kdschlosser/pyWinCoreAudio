
from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_DhtApi = CLSID('{30543F69-1213-4284-9D0A-5CE5F71A82CD}')
IID_IDhtApi = IID('{A2838F57-EB00-4F46-9567-A4C92AD7D730}')


class IDhtApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IDhtApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method Initialize')],
            HRESULT,
            'Initialize',
            (['in'], LPWSTR, 'pwszDeviceId')
        ),
        COMMETHOD(
            [comtypes.helpstring('method EnableDolbyHomeTheater')],
            HRESULT,
            'EnableDolbyHomeTheater',
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetDolbyHomeTheater')],
            HRESULT,
            'GetDolbyHomeTheater',
            (['out'], POINTER(LONG), 'pfEnable')
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
        COMMETHOD(
            [comtypes.helpstring('method GetHeadphoneMode')],
            HRESULT,
            'GetHeadphoneMode',
            (['out'], POINTER(INT), 'pfHeadphoneMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsEnableDolbyHomeTheaterChanged')],
            HRESULT,
            'IsEnableDolbyHomeTheaterChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(INT), 'pfEnabled')
        ),
    ]


class DhtApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_DhtApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IDhtApi]

    @staticmethod
    def IDhtApi() -> IDhtApi:
        return comtypes.CoCreateInstance(
            CLSID_DhtApi,
            IDhtApi,
            comtypes.CLSCTX_ALL
        )
