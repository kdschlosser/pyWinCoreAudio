from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib

CLSID_DPCEE3sApi = CLSID('{E0760680-E3E3-41A6-A5BE-275F5C21BDD9}')
IID_IDPCEE3sApi = IID('{9A17F557-4A5C-4907-9F0B-436C4E36F1EC}')
IID_IDPCEE3sApi2 = IID('{9257D0BF-35AA-49A1-B8E3-2847EABD2A90}')
IID_ISRSPSApi = IID('{C96F6BA8-B1BF-4752-AE53-4413118B2148}')


class IDPCEE3sApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IDPCEE3sApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method SetCurrentMode')],
            HRESULT,
            'SetCurrentMode',
            (['in'], INT, 'nMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetCurrentMode')],
            HRESULT,
            'GetCurrentMode',
            (['out', 'retval'], POINTER(INT), 'pnMode')
        ),
    ]


class IDPCEE3sApi2(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IDPCEE3sApi2
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method SetCurrentMode')],
            HRESULT,
            'SetCurrentMode',
            (['in'], LPWSTR, 'pwstrId'),
            (['in'], INT, 'nMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetCurrentMode')],
            HRESULT,
            'GetCurrentMode',
            (['in'], LPWSTR, 'pwstrId'),
            (['out', 'retval'], POINTER(INT), 'pnMode')
        ),
    ]


class ISRSPSApi(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_ISRSPSApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method SetEnable')],
            HRESULT,
            'SetEnable',
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetEnable')],
            HRESULT,
            'GetEnable',
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
    ]


class DPCEE3sApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_DPCEE3sApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IDPCEE3sApi, IDPCEE3sApi2, ISRSPSApi]

    @staticmethod
    def IDPCEE3sApi() -> IDPCEE3sApi:
        return comtypes.CoCreateInstance(
            CLSID_DPCEE3sApi,
            IDPCEE3sApi,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IDPCEE3sApi2() -> IDPCEE3sApi2:
        return comtypes.CoCreateInstance(
            CLSID_DPCEE3sApi,
            IDPCEE3sApi2,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def ISRSPSApi() -> ISRSPSApi:
        return comtypes.CoCreateInstance(
            CLSID_DPCEE3sApi,
            ISRSPSApi,
            comtypes.CLSCTX_ALL
        )
