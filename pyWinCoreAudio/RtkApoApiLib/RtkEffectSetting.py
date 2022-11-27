from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_RtkEffectSetting = CLSID('{C51A2350-6AB3-44A5-8B32-61A9DBEBCF45}')
IID_IRtkEffectSetting = IID('{1A00D230-51D2-4006-9AA8-3CBC09FDA649}')
IID_IRtkEffectSetting2 = IID('{E83FF7E2-18F6-4BAC-B762-BFF29FD3AB1C}')


class IRtkEffectSetting(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IRtkEffectSetting
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method SetCondition')],
            HRESULT,
            'SetCondition',
            (['in'], LPWSTR, 'pwszDeviceId'),
            (['in'], INT, 'nCondCnt'),
            (['in'], POINTER(GUID), 'pCondGuids'),
            (['in'], POINTER(INT), 'pCondValues')
        ),
        COMMETHOD(
            [comtypes.helpstring('method EnumCondition')],
            HRESULT,
            'EnumCondition',
            (['in'], LPWSTR, 'pwszDeviceId'),
            (['in'], POINTER(INT), 'pnCondCnt'),
            (['in'], POINTER(POINTER(GUID)), 'ppCondGuids')
        ),
    ]


class IRtkEffectSetting2(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IRtkEffectSetting2
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetCondDefs')],
            HRESULT,
            'GetCondDefs',
            (['in'], LPWSTR, 'pwszDeviceId'),
            (['in', 'out'], POINTER(INT), 'pnCount'),
            (['in', 'out'], POINTER(POINTER(INT)), 'ppCondCounts'),
            (['in', 'out'], POINTER(POINTER(POINTER(GUID))), 'pppCondGuids'),
            (['in', 'out'], POINTER(POINTER(POINTER(INT))), 'pppCondValues')
        ),
    ]


class RtkEffectSetting(comtypes.CoClass):
    _reg_clsid_ = CLSID_RtkEffectSetting
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IRtkEffectSetting, IRtkEffectSetting2]

    @staticmethod
    def IRtkEffectSetting() -> IRtkEffectSetting:
        return comtypes.CoCreateInstance(
            CLSID_RtkEffectSetting,
            IRtkEffectSetting,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IRtkEffectSetting2() -> IRtkEffectSetting2:
        return comtypes.CoCreateInstance(
            CLSID_RtkEffectSetting,
            IRtkEffectSetting2,
            comtypes.CLSCTX_ALL
        )
