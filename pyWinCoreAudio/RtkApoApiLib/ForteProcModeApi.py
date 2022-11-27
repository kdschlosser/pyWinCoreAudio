from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


IID_IForteProcModeApi = IID('{EC42102D-19E7-4938-820C-0B9D3748417C}')
CLSID_ForteProcModeApi = CLSID('{073E187F-BB72-4D0F-8368-C6A20CD9EA50}')


class IForteProcModeApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IForteProcModeApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetEndpointCount')],
            HRESULT,
            'GetEndpointCount',
            (['out', 'retval'], POINTER(INT), 'pcnt')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetModePKEYInfo')],
            HRESULT,
            'GetModePKEYInfo',
            (['in'], INT, 'nEndpointIndex'),
            (['out'], POINTER(WSTRING), 'psRegPath')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetProcMode')],
            HRESULT,
            'SetProcMode',
            (['in'], INT, 'nEndpointIndex'),
            (['in'], INT, 'nMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetProcMode')],
            HRESULT,
            'GetProcMode',
            (['in'], INT, 'nEndpointIndex'),
            (['out', 'retval'], POINTER(INT), 'pnMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetBeamForming')],
            HRESULT,
            'SetBeamForming',
            (['in'], INT, 'nEndpointIndex'),
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetBeamForming')],
            HRESULT,
            'GetBeamForming',
            (['in'], INT, 'nEndpointIndex'),
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetFarFieldPickup')],
            HRESULT,
            'SetFarFieldPickup',
            (['in'], INT, 'nEndpointIndex'),
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetFarFieldPickup')],
            HRESULT,
            'GetFarFieldPickup',
            (['in'], INT, 'nEndpointIndex'),
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetAEC')],
            HRESULT,
            'SetAEC',
            (['in'], INT, 'nEndpointIndex'),
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetAEC')],
            HRESULT,
            'GetAEC',
            (['in'], INT, 'nEndpointIndex'),
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetKeyStroke')],
            HRESULT,
            'SetKeyStroke',
            (['in'], INT, 'nEndpointIndex'),
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetKeyStroke')],
            HRESULT,
            'GetKeyStroke',
            (['in'], INT, 'nEndpointIndex'),
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetForteDisableAll')],
            HRESULT,
            'SetForteDisableAll',
            (['in'], INT, 'nEndpointIndex'),
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetForteDisableAll')],
            HRESULT,
            'GetForteDisableAll',
            (['in'], INT, 'nEndpointIndex'),
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetForteVRMode')],
            HRESULT,
            'SetForteVRMode',
            (['in'], INT, 'nEndpointIndex'),
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetForteVRMode')],
            HRESULT,
            'GetForteVRMode',
            (['in'], INT, 'nEndpointIndex'),
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
    ]


class ForteProcModeApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_ForteProcModeApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IForteProcModeApi]

    @staticmethod
    def IForteProcModeApi() -> IForteProcModeApi:
        return comtypes.CoCreateInstance(
            CLSID_ForteProcModeApi,
            IForteProcModeApi,
            comtypes.CLSCTX_ALL
        )
