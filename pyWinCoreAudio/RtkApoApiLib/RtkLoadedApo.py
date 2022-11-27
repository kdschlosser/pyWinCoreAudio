from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib

CLSID_RtkLoadedApo = CLSID('{FBEBFC99-D58A-48F2-BD9B-96A9DE7DBDBA}')
IID_IRtkLoadedApo = IID('{A5E23FED-11B1-4C0F-9392-8D44E1EE8F51}')


class IRtkLoadedApo(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IRtkLoadedApo
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetLoadedApoGUID')],
            HRESULT,
            'GetLoadedApoGUID',
            (['in'], LPWSTR, 'pwszDeviceId'),
            (['in'], INT, 'nRoleBits'),
            (['out'], POINTER(POINTER(GUID)), 'ppAposIds'),
            (['out'], POINTER(UINT), 'pcApos')
        ),
    ]


class RtkLoadedApo(comtypes.CoClass):
    _reg_clsid_ = CLSID_RtkLoadedApo
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IRtkLoadedApo]

    @staticmethod
    def IRtkLoadedApo() -> IRtkLoadedApo:
        return comtypes.CoCreateInstance(
            CLSID_RtkLoadedApo,
            IRtkLoadedApo,
            comtypes.CLSCTX_ALL
        )
