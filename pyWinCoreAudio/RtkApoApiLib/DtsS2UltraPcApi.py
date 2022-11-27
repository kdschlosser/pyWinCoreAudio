from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_DtsS2UltraPcApi = CLSID('{9C4B415C-B076-43F4-928D-98D8B465087A}')
IID_IDtsS2UltraPcApi = IID('{3FFE1886-991C-423F-B3A7-E0E2FE8CCF40}')


class IDtsS2UltraPcApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IDtsS2UltraPcApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method Initialize')],
            HRESULT,
            'Initialize',
            (['in'], LPWSTR, 'pwszDeviceId')
        ),
        COMMETHOD(
            [comtypes.helpstring('method EnableS2UltraPc')],
            HRESULT,
            'EnableS2UltraPc',
            (['in'], LONG, 'bEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetS2UltraPc')],
            HRESULT,
            'GetS2UltraPc',
            (['out', 'retval'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetNeoPcMode')],
            HRESULT,
            'SetNeoPcMode',
            (['in'], INT, 'nNeoMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetNeoPcMode')],
            HRESULT,
            'GetNeoPcMode',
            (['out', 'retval'], POINTER(INT), 'pnNeoMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method EnableBassEnhancement')],
            HRESULT,
            'EnableBassEnhancement',
            (['in'], LONG, 'bEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetBassEnhancement')],
            HRESULT,
            'GetBassEnhancement',
            (['out', 'retval'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetBassEnhancementGain')],
            HRESULT,
            'SetBassEnhancementGain',
            (['in'], INT, 'nGain')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetBassEnhancementGain')],
            HRESULT,
            'GetBassEnhancementGain',
            (['out', 'retval'], POINTER(INT), 'pnGain')
        ),
        COMMETHOD(
            [comtypes.helpstring('method EnableDialogClarity')],
            HRESULT,
            'EnableDialogClarity',
            (['in'], LONG, 'bEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetDialogClarity')],
            HRESULT,
            'GetDialogClarity',
            (['out', 'retval'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetDialogClarityGain')],
            HRESULT,
            'SetDialogClarityGain',
            (['in'], INT, 'nGain')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetDialogClarityGain')],
            HRESULT,
            'GetDialogClarityGain',
            (['out', 'retval'], POINTER(INT), 'pnGain')
        ),
        COMMETHOD(
            [comtypes.helpstring('method EnableLfeMixing')],
            HRESULT,
            'EnableLfeMixing',
            (['in'], LONG, 'bEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetLfeMixing')],
            HRESULT,
            'GetLfeMixing',
            (['out', 'retval'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsEnableS2UltraPcChanged')],
            HRESULT,
            'IsEnableS2UltraPcChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsNeoPcModeChanged')],
            HRESULT,
            'IsNeoPcModeChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(INT), 'pnNeoMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsEnableBassEnhancementChanged')],
            HRESULT,
            'IsEnableBassEnhancementChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsBassEnhancementGainChanged')],
            HRESULT,
            'IsBassEnhancementGainChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(INT), 'pnGain')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsEnableDialogClarityChanged')],
            HRESULT,
            'IsEnableDialogClarityChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsDialogClarityGainChanged')],
            HRESULT,
            'IsDialogClarityGainChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(INT), 'pnGain')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsEnableLfeMixingChanged')],
            HRESULT,
            'IsEnableLfeMixingChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(LONG), 'pbEnabled')
        ),
    ]


class DtsS2UltraPcApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_DtsS2UltraPcApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IDtsS2UltraPcApi]

    @staticmethod
    def IDtsS2UltraPcApi() -> IDtsS2UltraPcApi:
        return comtypes.CoCreateInstance(
            CLSID_DtsS2UltraPcApi,
            IDtsS2UltraPcApi,
            comtypes.CLSCTX_ALL
        )
