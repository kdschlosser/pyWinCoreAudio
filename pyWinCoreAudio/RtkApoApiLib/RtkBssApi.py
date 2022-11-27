from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib

CLSID_RtkBssApi = CLSID('{2438FE00-BD99-49AF-9E48-D3B03885FA8B}')
IID_IRtkBssApi = IID('{A60A9DD6-410A-49AC-9F43-9C6FD355E138}')
IID_IRtkBssApi2 = IID('{A5564785-34A2-4A03-BEA9-0065E4DCFA5A}')
IID_IRtkBssApi3 = IID('{FD5DA846-3248-4FF5-8C93-B10E6A2894B6}')
IID_IRtkBssApi4 = IID('{AC1F6C3D-EB3A-48AE-A999-F847B4FB6383}')


class IRtkBssApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IRtkBssApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method Initialize')],
            HRESULT,
            'Initialize',
            (['in'], LPWSTR, 'pwszDeviceId')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetBssMode')],
            HRESULT,
            'SetBssMode',
            (['in'], INT, 'nMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetBssMode')],
            HRESULT,
            'GetBssMode',
            (['out'], POINTER(INT), 'pnMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetAec')],
            HRESULT,
            'SetAEC',
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetAec')],
            HRESULT,
            'GetAEC',
            (['out'], POINTER(LONG), 'pfEnable')
        ),
    ]


class IRtkBssApi2(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IRtkBssApi2
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method IsGen3Supported')],
            HRESULT,
            'IsGen3Supported',
            (['out'], POINTER(LONG), 'pfSupported')
        ),
    ]


class IRtkBssApi3(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IRtkBssApi3
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method IsGen2Supported')],
            HRESULT,
            'IsGen2Supported',
            (['out'], POINTER(LONG), 'pfSupported')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsGen3ForAlexaSupported')],
            HRESULT,
            'IsGen3ForAlexaSupported',
            (['out'], POINTER(LONG), 'pfSupported')
        ),
    ]


class IRtkBssApi4(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IRtkBssApi4
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetRenderNSState')],
            HRESULT,
            'GetRenderNSState',
            (['out'], POINTER(INT), 'pfAuto'),
            (['out'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetRenderNSState')],
            HRESULT,
            'SetRenderNSState',
            (['in'], INT, 'fAuto'),
            (['in'], LONG, 'fEnable')
        ),
    ]


class RtkBssApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_RtkBssApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [
        IRtkBssApi,
        IRtkBssApi2,
        IRtkBssApi3,
        IRtkBssApi4
    ]

    @staticmethod
    def IRtkBssApi() -> IRtkBssApi:
        return comtypes.CoCreateInstance(
            CLSID_RtkBssApi,
            IRtkBssApi,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IRtkBssApi2() -> IRtkBssApi2:
        return comtypes.CoCreateInstance(
            CLSID_RtkBssApi,
            IRtkBssApi2,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IRtkBssApi3() -> IRtkBssApi3:
        return comtypes.CoCreateInstance(
            CLSID_RtkBssApi,
            IRtkBssApi3,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IRtkBssApi4() -> IRtkBssApi4:
        return comtypes.CoCreateInstance(
            CLSID_RtkBssApi,
            IRtkBssApi4,
            comtypes.CLSCTX_ALL
        )
