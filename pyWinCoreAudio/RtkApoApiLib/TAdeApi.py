from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib

CLSID_TAdeApi = CLSID('{45750165-6E3D-4413-AEE0-D486BC169E0B}')
IID_ITAdeApi = IID('{E67F1B19-D3EA-4469-8A0F-EE17BEEA6BD6}')
IID_ITEEApi = IID('{6027D630-511E-48BC-A815-20D154550105}')


class ITAdeApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_ITAdeApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method SetAde')],
            HRESULT,
            'SetAde',
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetAde')],
            HRESULT,
            'GetAde',
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
    ]

    @property
    def Ade(self) -> bool:
        return bool(self.GetEE())

    @Ade.setter
    def Ade(self, value: bool):
        if value:
            value = 1
        else:
            value = 0

        self.SetEE(value)


class ITEEApi(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_ITEEApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method SetEE')],
            HRESULT,
            'SetEE',
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetEE')],
            HRESULT,
            'GetEE',
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
    ]

    @property
    def EE(self) -> bool:
        return bool(self.GetEE())

    @EE.setter
    def EE(self, value: bool):
        if value:
            value = 1
        else:
            value = 0

        self.SetEE(value)



class TAdeApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_TAdeApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [ITAdeApi, ITEEApi]

    @staticmethod
    def ITAdeApi() -> ITAdeApi:
        return comtypes.CoCreateInstance(
            CLSID_TAdeApi,
            ITAdeApi,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def ITEEApi() -> ITEEApi:
        return comtypes.CoCreateInstance(
            CLSID_TAdeApi,
            ITEEApi,
            comtypes.CLSCTX_ALL
        )
