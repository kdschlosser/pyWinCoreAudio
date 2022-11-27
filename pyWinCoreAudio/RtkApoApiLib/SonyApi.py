from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib
from ..propidl import PROPVARIANT


CLSID_SonyApi = CLSID('{C26C3384-7112-4B5A-8BB5-1328D9543D79}')
IID_ISonyApi = IID('{A90FAE67-C02E-45E1-81C6-E6082552784E}')
IID_ISonyApi2 = IID('{51FDA9E7-DAC0-4722-8BC5-376FFB8BD69F}')


class ISonyApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_ISonyApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method SetRealtekFeatures')],
            HRESULT,
            'SetRealtekFeatures',
            (['in'], LPWSTR, 'pwstrId'),
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetRealtekFeatures')],
            HRESULT,
            'GetRealtekFeatures',
            (['in'], LPWSTR, 'pwstrId'),
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetSonyFeatures')],
            HRESULT,
            'SetSonyFeatures',
            (['in'], LPWSTR, 'pwstrId'),
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSonyFeatures')],
            HRESULT,
            'GetSonyFeatures',
            (['in'], LPWSTR, 'pwstrId'),
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
    ]


class ISonyApi2(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_ISonyApi2
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method SetSony_78D56A5E')],
            HRESULT,
            'SetSony_78D56A5E',
            (['in'], LPWSTR, 'pwstrId'),
            (['in'], ULONG, 'pid'),
            (['in'], POINTER(PROPVARIANT), 'pv')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSony_78D56A5E')],
            HRESULT,
            'GetSony_78D56A5E',
            (['in'], LPWSTR, 'pwstrId'),
            (['in'], ULONG, 'pid'),
            (['in', 'out'], POINTER(PROPVARIANT), 'pv')
        ),
    ]


class SonyApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_SonyApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [ISonyApi, ISonyApi2]

    @staticmethod
    def ISonyApi() -> ISonyApi:
        return comtypes.CoCreateInstance(
            CLSID_SonyApi,
            ISonyApi,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def ISonyApi2() -> ISonyApi2:
        return comtypes.CoCreateInstance(
            CLSID_SonyApi,
            ISonyApi2,
            comtypes.CLSCTX_ALL
        )

