from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_SAApi = CLSID('{38361248-B055-4EA8-A373-940464FB318E}')
IID_ISAApi = IID('{03C81CEF-DA23-49AB-8BFA-E35CE08AE178}')
IID_ISAApi2 = IID('{6C431600-3688-46BA-962B-353342AB3530}')


class ISAApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_ISAApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method SetSAParameters')],
            HRESULT,
            'SetSAParameters',
            (['in'], POINTER(UBYTE), 'pData'),
            (['in'], ULONG, 'cbSize')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSAParameters')],
            HRESULT,
            'GetSAParameters',
            (['out'], POINTER(POINTER(UBYTE)), 'ppData'),
            (['out', 'retval'], POINTER(ULONG), 'pcbSize')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetSAUserParameters')],
            HRESULT,
            'SetSAUserParameters',
            (['in'], POINTER(UBYTE), 'pData'),
            (['in'], ULONG, 'cbSize')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSAUserParameters')],
            HRESULT,
            'GetSAUserParameters',
            (['out'], POINTER(POINTER(UBYTE)), 'ppData'),
            (['out', 'retval'], POINTER(ULONG), 'pcbSize')
        ),
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
        COMMETHOD(
            [comtypes.helpstring('method SetPreset')],
            HRESULT,
            'SetPreset',
            (['in'], INT, 'nPresetIndex')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetPreset')],
            HRESULT,
            'GetPreset',
            (['out', 'retval'], POINTER(INT), 'pnPresetIndex')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetSAAmbient')],
            HRESULT,
            'SetSAAmbient',
            (['in'], INT, 'nContentType'),
            (['in'], INT, 'nAmbient')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSAAmbient')],
            HRESULT,
            'GetSAAmbient',
            (['in'], INT, 'nContentType'),
            (['out', 'retval'], POINTER(INT), 'pnAmbient')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetSACutoff')],
            HRESULT,
            'SetSACutoff',
            (['in'], INT, 'nLeft'),
            (['in'], INT, 'nRight')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSACutoff')],
            HRESULT,
            'GetSACutoff',
            (['in'], POINTER(INT), 'pnLeft'),
            (['in'], POINTER(INT), 'pnRight')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetSAAutoVolume')],
            HRESULT,
            'SetSAAutoVolume',
            (['in'], INT, 'nContentType'),
            (['in'], INT, 'nPresetIndex'),
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSAAutoVolume')],
            HRESULT,
            'GetSAAutoVolume',
            (['in'], INT, 'nContentType'),
            (['in'], INT, 'nPresetIndex'),
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetSAUserEQ')],
            HRESULT,
            'SetSAUserEQ',
            (['in'], INT, 'nPresetIndex'),
            (['in'], INT, 'nBandCnt'),
            (['in'], POINTER(INT), 'pnBandData')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSAUserEQ')],
            HRESULT,
            'GetSAUserEQ',
            (['in'], INT, 'nPresetIndex'),
            (['in'], INT, 'nBandCnt'),
            (['out', 'retval'], POINTER(INT), 'pnBandData')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetSAUser3D')],
            HRESULT,
            'SetSAUser3D',
            (['in'], INT, 'nPresetIndex'),
            (['in'], INT, 'n3DLevel')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSAUser3D')],
            HRESULT,
            'GetSAUser3D',
            (['in'], INT, 'nPresetIndex'),
            (['out', 'retval'], POINTER(INT), 'pn3DLevel')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetSAUserBE')],
            HRESULT,
            'SetSAUserBE',
            (['in'], INT, 'nPresetIndex'),
            (['in'], INT, 'nBELevel')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSAUserBE')],
            HRESULT,
            'GetSAUserBE',
            (['in'], INT, 'nPresetIndex'),
            (['out', 'retval'], POINTER(INT), 'pnBELevel')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetSAUserCH')],
            HRESULT,
            'SetSAUserCH',
            (['in'], INT, 'nPresetIndex'),
            (['in'], INT, 'nRoomSize'),
            (['in'], INT, 'nLevel')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSAUserCH')],
            HRESULT,
            'GetSAUserCH',
            (['in'], INT, 'nPresetIndex'),
            (['out'], POINTER(INT), 'pnRoomSize'),
            (['out', 'retval'], POINTER(INT), 'pnLevel')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetSAUserCla')],
            HRESULT,
            'SetSAUserCla',
            (['in'], INT, 'nPresetIndex'),
            (['in'], INT, 'nClaLevel')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSAUserCla')],
            HRESULT,
            'GetSAUserCla',
            (['in'], INT, 'nPresetIndex'),
            (['out', 'retval'], POINTER(INT), 'pnClaLevel')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetSAUserAuUp')],
            HRESULT,
            'SetSAUserAuUp',
            (['in'], INT, 'nPresetIndex'),
            (['in'], INT, 'nAuUpLevel')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSAUserAuUp')],
            HRESULT,
            'GetSAUserAuUp',
            (['in'], INT, 'nPresetIndex'),
            (['out', 'retval'], POINTER(INT), 'pnAuUpLevel')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetDMCSmartVolume')],
            HRESULT,
            'SetDMCSmartVolume',
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetDMCSmartVolume')],
            HRESULT,
            'GetDMCSmartVolume',
            (['out', 'retval'], POINTER(LONG), 'pfEnable')
        ),
    ]


class ISAApi2(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_ISAApi2
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method SetDMC3DDepth')],
            HRESULT,
            'SetDMC3DDepth',
            (['in'], INT, 'nDepth')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetDMC3DDepth')],
            HRESULT,
            'GetDMC3DDepth',
            (['out', 'retval'], POINTER(INT), 'pnDepth')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetDMC3DDepthFreqGain')],
            HRESULT,
            'SetDMC3DDepthFreqGain',
            (['in'], INT, 'nSpkF0L'),
            (['in'], INT, 'nSpkF0R'),
            (['in'], INT, 'nLastGL100'),
            (['in'], INT, 'nLastGR100')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetDMC3DDepthFreqGain')],
            HRESULT,
            'GetDMC3DDepthFreqGain',
            (['out'], POINTER(INT), 'pnSpkF0L'),
            (['out'], POINTER(INT), 'pnSpkF0R'),
            (['out'], POINTER(INT), 'pnLastGL100'),
            (['out'], POINTER(INT), 'pnLastGR100')
        ),
    ]


class SAApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_SAApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [ISAApi, ISAApi2]

    @staticmethod
    def ISAApi() -> ISAApi:
        return comtypes.CoCreateInstance(
            CLSID_SAApi,
            ISAApi,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def ISAApi2() -> ISAApi2:
        return comtypes.CoCreateInstance(
            CLSID_SAApi,
            ISAApi2,
            comtypes.CLSCTX_ALL
        )
