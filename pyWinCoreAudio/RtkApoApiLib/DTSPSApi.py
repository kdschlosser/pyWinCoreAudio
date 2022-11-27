from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_DTSPSApi = CLSID('{CE93F536-D2FB-4E8F-8193-F57AD8BC90D5}')
IID_IDTSPSApi = IID('{277978B5-A254-4246-B4C9-766596C83810}')


class IDTSPSApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IDTSPSApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method SetEndpointId')],
            HRESULT,
            'SetEndpointId',
            (['in'], LPWSTR, 'pwszDeviceId')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetEndpointId')],
            HRESULT,
            'GetEndpointId',
            (['out', 'retval'], POINTER(LPWSTR), 'ppwszDeviceId')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSpeakerMode')],
            HRESULT,
            'GetSpeakerMode',
            (['out', 'retval'], POINTER(INT), 'pnSpkrMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method ResetSurroundSensation')],
            HRESULT,
            'ResetSurroundSensation',
            (['in'], INT, 'nSpkrMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetSurroundSensation')],
            HRESULT,
            'SetSurroundSensation',
            (['in'], INT, 'nSpkrMode'),
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSurroundSensation')],
            HRESULT,
            'GetSurroundSensation',
            (['in'], INT, 'nSpkrMode'),
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetSSContentType')],
            HRESULT,
            'SetSSContentType',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSSContentType')],
            HRESULT,
            'GetSSContentType',
            (['in'], INT, 'nSpkrMode'),
            (['out', 'retval'], POINTER(INT), 'pnContentType')
        ),
        COMMETHOD(
            [comtypes.helpstring('method ResetBoost')],
            HRESULT,
            'ResetBoost'
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetBoost')],
            HRESULT,
            'SetBoost',
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetBoost')],
            HRESULT,
            'GetBoost',
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
    ]


class DTSPSApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_DTSPSApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IDTSPSApi]

    @staticmethod
    def IDTSPSApi() -> IDTSPSApi:
        return comtypes.CoCreateInstance(
            CLSID_DTSPSApi,
            IDTSPSApi,
            comtypes.CLSCTX_ALL
        )

