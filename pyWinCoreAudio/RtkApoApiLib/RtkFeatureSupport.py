from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_RtkFeatureSupport = CLSID('{3B8ED366-08BE-4C3E-98C2-324F45AE0FF8}')
IID_IRtkFeatureSupport = IID('{7E2BF545-78C4-4D8C-BC56-8F8CC293BD0C}')


class IRtkFeatureSupport(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IRtkFeatureSupport
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method Initialize')],
            HRESULT,
            'QueryFeatureSupport',
            (['in'], LPWSTR, 'pwszDeviceId'),
            (['in'], INT, 'nFeatureId'),
            (['out'], POINTER(INT), 'pnSupportCode')
        ),
    ]


class RtkFeatureSupport(comtypes.CoClass):
    _reg_clsid_ = CLSID_RtkFeatureSupport
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IRtkFeatureSupport]

    @staticmethod
    def IRtkFeatureSupport() -> IRtkFeatureSupport:
        return comtypes.CoCreateInstance(
            CLSID_RtkFeatureSupport,
            IRtkFeatureSupport,
            comtypes.CLSCTX_ALL
        )
