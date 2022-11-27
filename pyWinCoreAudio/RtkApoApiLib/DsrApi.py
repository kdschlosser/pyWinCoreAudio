from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_DsrApi = CLSID('{048E5288-25FF-4552-BFE0-A9B7DE629CBA}')
IID_IDsrApi = IID('{BDB15C48-B756-4463-BAAD-7935DC57FA7E}')


class IDsrApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IDsrApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method Initialize')],
            HRESULT,
            'Initialize',
            (['in'], LPWSTR, 'pwszDeviceId')
        ),
        COMMETHOD(
            [comtypes.helpstring('method EnableDolbySoundRoom')],
            HRESULT,
            'EnableDolbySoundRoom',
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetDolbySoundRoom')],
            HRESULT,
            'GetDolbySoundRoom',
            (['out'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetHeadphoneMode')],
            HRESULT,
            'GetHeadphoneMode',
            (['out'], POINTER(INT), 'pfHeadphoneMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsEnableDolbySoundRoomChanged')],
            HRESULT,
            'IsEnableDolbySoundRoomChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(INT), 'pfEnabled')
        ),
    ]


class DsrApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_DsrApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IDsrApi]

    @staticmethod
    def IDsrApi() -> IDsrApi:
        return comtypes.CoCreateInstance(
            CLSID_DsrApi,
            IDsrApi,
            comtypes.CLSCTX_ALL
        )
