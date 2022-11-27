from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_DTSU2PApi = CLSID('{42D0571B-9239-4EEE-B8A0-1C8B85608894}')

IID_IDTSU2PApi = IID('{BFC59E45-FDAE-4DE1-A47A-9DB1C64D5AB4}')


class IDTSU2PApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IDTSU2PApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method Reset')],
            HRESULT,
            'Reset'
        ),
        COMMETHOD(
            [comtypes.dispid(2)],
            HRESULT,
            'GetDtsMasterEnabled',
            (['out', 'retval'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(3)],
            HRESULT,
            'SetDtsMasterEnabled',
            (['in'], LONG, 'bEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(4)],
            HRESULT,
            'GetDtsS2Enabled',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['out', 'retval'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(5)],
            HRESULT,
            'SetDtsS2Enabled',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['in'], LONG, 'bEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(6)],
            HRESULT,
            'GetVoiceClarityGain',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['out', 'retval'], POINTER(FLOAT), 'pfGain')
        ),
        COMMETHOD(
            [comtypes.dispid(7)],
            HRESULT,
            'SetVoiceClarityGain',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['in'], FLOAT, 'fGain')
        ),
        COMMETHOD(
            [comtypes.dispid(8)],
            HRESULT,
            'GetBassEnhancementEnabled',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['out', 'retval'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(9)],
            HRESULT,
            'SetBassEnhancementEnabled',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['in'], LONG, 'bEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(10)],
            HRESULT,
            'GetBassGain',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['out', 'retval'], POINTER(FLOAT), 'pfGain')
        ),
        COMMETHOD(
            [comtypes.dispid(11)],
            HRESULT,
            'SetBassGain',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['in'], FLOAT, 'fGain')
        ),
        COMMETHOD(
            [comtypes.dispid(12)],
            HRESULT,
            'GetSymmetryEnabled',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['out', 'retval'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(13)],
            HRESULT,
            'SetSymmetryEnabled',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['in'], LONG, 'bEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(14)],
            HRESULT,
            'GetBoostEnabled',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['out', 'retval'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(15)],
            HRESULT,
            'SetBoostEnabled',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['in'], LONG, 'bEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(16)],
            HRESULT,
            'GetBoostLevel',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['out', 'retval'], POINTER(ULONG), 'pulLevel')
        ),
        COMMETHOD(
            [comtypes.dispid(17)],
            HRESULT,
            'SetBoostLevel',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['in'], ULONG, 'ulLevel')
        ),
        COMMETHOD(
            [comtypes.dispid(18)],
            HRESULT,
            'GetLeqEnabled',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['out', 'retval'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(19)],
            HRESULT,
            'SetLeqEnabled',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['in'], LONG, 'bEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(20)],
            HRESULT,
            'GetHFCompFactor',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['out', 'retval'], POINTER(FLOAT), 'pfFactor')
        ),
        COMMETHOD(
            [comtypes.dispid(21)],
            HRESULT,
            'SetHFCompFactor',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['in'], FLOAT, 'fFactor')
        ),
        COMMETHOD(
            [comtypes.dispid(22)],
            HRESULT,
            'GetLFCompFactor',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['out', 'retval'], POINTER(FLOAT), 'pfFactor')
        ),
        COMMETHOD(
            [comtypes.dispid(23)],
            HRESULT,
            'SetLFCompFactor',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['in'], FLOAT, 'fFactor')
        ),
        COMMETHOD(
            [comtypes.dispid(24)],
            HRESULT,
            'GetEnvncEnabled',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['out', 'retval'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(25)],
            HRESULT,
            'SetClearAudioEnabled',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['in'], LONG, 'bEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(26)],
            HRESULT,
            'GetEnvcCompFactor',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['out', 'retval'], POINTER(FLOAT), 'pfPower')
        ),
        COMMETHOD(
            [comtypes.dispid(27)],
            HRESULT,
            'SetEnvcCompFactor',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType'),
            (['in'], FLOAT, 'fPower')
        ),
        COMMETHOD(
            [comtypes.dispid(28)],
            HRESULT,
            'GetMAPEnabled',
            (['out', 'retval'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(29)],
            HRESULT,
            'SetMAPEnabled',
            (['in'], LONG, 'bEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(30)],
            HRESULT,
            'GetAESEnabled',
            (['out', 'retval'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(31)],
            HRESULT,
            'SetAESEnabled',
            (['in'], LONG, 'bEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(32)],
            HRESULT,
            'GetBFEnabled',
            (['out', 'retval'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.dispid(33)],
            HRESULT,
            'SetBFEnabled',
            (['in'], LONG, 'bEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetEndpointId')],
            HRESULT,
            'GetEndpointId',
            (['out', 'retval'], POINTER(LPWSTR), 'ppwszDeviceId')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetEndpointId')],
            HRESULT,
            'SetEndpointId',
            (['in'], LPWSTR, 'pwszDeviceId')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSpeakerMode')],
            HRESULT,
            'GetSpeakerMode',
            (['out', 'retval'], POINTER(INT), 'pnSpkrMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetContentType')],
            HRESULT,
            'GetContentType',
            (['in'], INT, 'nSpkrMode'),
            (['out', 'retval'], POINTER(INT), 'pnContentType')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetContentType')],
            HRESULT,
            'SetContentType',
            (['in'], INT, 'nSpkrMode'),
            (['in'], INT, 'nContentType')
        ),
    ]


class DTSU2PApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_DTSU2PApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IDTSU2PApi]

    @staticmethod
    def IDTSU2PApi() -> IDTSU2PApi:
        return comtypes.CoCreateInstance(
            CLSID_DTSU2PApi,
            IDTSU2PApi,
            comtypes.CLSCTX_ALL
        )
