from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_DSpruseApi = CLSID('{9914A841-B21A-489C-949C-89306F33BF9A}')
IID_IDSpruseApi = IID('{500AE005-8E20-4781-BBF9-96A640AB421A}')


class IDSpruseApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IDSpruseApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method Initialize')],
            HRESULT,
            'Initialize',
            (['in'], LPWSTR, 'pwszDeviceId')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetPceeSku')],
            HRESULT,
            'GetPceeSku',
            (['out', 'retval'], POINTER(ULONG), 'pdwPceeSku')
        ),
        COMMETHOD(
            [comtypes.helpstring('method EnableSoundSpaceExpander')],
            HRESULT,
            'EnableSoundSpaceExpander',
            (['in'], LONG, 'bEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetSoundSpaceExpander')],
            HRESULT,
            'GetSoundSpaceExpander',
            (['out', 'retval'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method EnableNatureBass')],
            HRESULT,
            'EnableNatureBass',
            (['in'], LONG, 'bEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetNatureBass')],
            HRESULT,
            'GetNatureBass',
            (['out', 'retval'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method EnableDolbyHeadphone')],
            HRESULT,
            'EnableDolbyHeadphone',
            (['in'], LONG, 'bEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetDolbyHeadphone')],
            HRESULT,
            'GetDolbyHeadphone',
            (['out', 'retval'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method EnablePrologicIIx')],
            HRESULT,
            'EnablePrologicIIx',
            (['in'], LONG, 'bEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetPrologicIIx')],
            HRESULT,
            'GetPrologicIIx',
            (['out', 'retval'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetCenterWidth')],
            HRESULT,
            'SetCenterWidth',
            (['in'], ULONG, 'dwCenterWidth')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetCenterWidth')],
            HRESULT,
            'GetCenterWidth',
            (['out', 'retval'], POINTER(ULONG), 'pdwCenterWidth')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetDimension')],
            HRESULT,
            'SetDimension',
            (['in'], ULONG, 'dwDimension')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetDimension')],
            HRESULT,
            'GetDimension',
            (['out', 'retval'], POINTER(ULONG), 'pdwDimension')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetNatureBassBoost')],
            HRESULT,
            'SetNatureBassBoost',
            (['in'], ULONG, 'nBoost')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetNatureBassBoost')],
            HRESULT,
            'GetNatureBassBoost',
            (['out', 'retval'], POINTER(ULONG), 'pnBoost')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsEnableSoundSpaceExpanderChanged')],
            HRESULT,
            'IsEnableSoundSpaceExpanderChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsEnableNatureBassChanged')],
            HRESULT,
            'IsEnableNatureBassChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsEnableDolbyHeadphoneChanged')],
            HRESULT,
            'IsEnableDolbyHeadphoneChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsEnablePrologicIIxChanged')],
            HRESULT,
            'IsEnablePrologicIIxChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(LONG), 'pbEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsCenterWidthChanged')],
            HRESULT,
            'IsCenterWidthChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(ULONG), 'pdwCenterWidth')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsDimensionChanged')],
            HRESULT,
            'IsDimensionChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(ULONG), 'pdwDimension')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsNatureBassBoostChanged')],
            HRESULT,
            'IsNatureBassBoostChanged',
            (['in'], PROPERTYKEY, 'key'),
            ([], POINTER(ULONG), 'pnBoost')
        ),
    ]


class DSpruseApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_DSpruseApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IDSpruseApi]

    @staticmethod
    def IDSpruseApi() -> IDSpruseApi:
        return comtypes.CoCreateInstance(
            CLSID_DSpruseApi,
            IDSpruseApi,
            comtypes.CLSCTX_ALL
        )
