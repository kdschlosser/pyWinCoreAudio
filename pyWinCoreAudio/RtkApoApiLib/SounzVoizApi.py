from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib

CLSID_SounzVoizApi = CLSID('{44EE5F71-5DD2-4434-B1AA-16B850F4BEC5}')
IID_ISounzVoizApiCallback = IID('{C7BF11C8-DA43-4761-BC7A-89BA67F44B09}')
IID_ISounzVoizApi = IID('{F2295F22-F237-4208-8483-653499E10BDA}')


class ISounzVoizApiCallback(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_ISounzVoizApiCallback
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [],
            HRESULT,
            'OnMasterEnableChange',
            (['in'], INT, 'bPlayback'),
            (['in'], INT, 'bEnable')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'OnActivePresetChange',
            (['in'], INT, 'bPlayback'),
            (['in'], INT, 'nPresetId')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'OnCustomTechValueChange',
            (['in'], INT, 'bPlayback'),
            (['in'], INT, 'nParamId'),
            (['in'], INT, 'nValue')
        ),
    ]


class ISounzVoizApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_ISounzVoizApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.dispid(1)],
            HRESULT,
            'GetMasterEnable',
            (['out', 'retval'], POINTER(INT), 'pbEnable')
        ),
        COMMETHOD(
            [comtypes.dispid(2)],
            HRESULT,
            'SetMasterEnable',
            (['in'], INT, 'bEnable')
        ),
        COMMETHOD(
            [comtypes.dispid(3)],
            HRESULT,
            'GetActivePreset',
            (['out'], POINTER(INT), 'pnPresetId'),
            (['out'], POINTER(INT), 'pnSpkrMode')
        ),
        COMMETHOD(
            [comtypes.dispid(4)],
            HRESULT,
            'SetActivePreset',
            (['in'], INT, 'nPresetId'),
            (['in'], INT, 'nSpkrMode')
        ),
        COMMETHOD(
            [comtypes.dispid(5)],
            HRESULT,
            'GetTechValue',
            (['in'], INT, 'nPresetId'),
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nParamId'),
            (['out'], POINTER(INT), 'pnValue')
        ),
        COMMETHOD(
            [comtypes.dispid(6)],
            HRESULT,
            'SetTechValue',
            (['in'], INT, 'nPresetId'),
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nParamId'),
            (['in'], INT, 'nValue')
        ),
        COMMETHOD(
            [comtypes.dispid(7)],
            HRESULT,
            'RegisterChangeNotifications',
            (['in'], POINTER(ISounzVoizApiCallback), 'pApiCallback')
        ),
        COMMETHOD(
            [comtypes.dispid(8)],
            HRESULT,
            'UnregisterChangeNotifications'
        ),
    ]


class SounzVoizApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_SounzVoizApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [ISounzVoizApi]

    @staticmethod
    def ISounzVoizApi() -> ISounzVoizApi:
        return comtypes.CoCreateInstance(
            CLSID_SounzVoizApi,
            ISounzVoizApi,
            comtypes.CLSCTX_ALL
        )
