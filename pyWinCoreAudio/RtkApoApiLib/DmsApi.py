
from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_DmsApi = CLSID('{B4928187-D2E3-4CFA-A2C2-A9C97B4956EA}')
IID_IDmsApi = IID('{E9B5E130-9157-43C5-8000-0EC6F5C90EBD}')


class IDmsApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IDmsApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method Initialize')],
            HRESULT,
            'Initialize',
            (['in'], LPWSTR, 'pwszDeviceId')
        ),
        COMMETHOD(
            [comtypes.helpstring('method EnableDolbyPrologic')],
            HRESULT,
            'EnableDolbyPrologic',
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetDolbyPrologic')],
            HRESULT,
            'GetDolbyPrologic',
            (['out'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method EnableDolbyVirtualSpeaker')],
            HRESULT,
            'EnableDolbyVirtualSpeaker',
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetDolbyVirtualSpeaker')],
            HRESULT,
            'GetDolbyVirtualSpeaker',
            (['out'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method EnableDolbyHeadphone')],
            HRESULT,
            'EnableDolbyHeadphone',
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetDolbyHeadphone')],
            HRESULT,
            'GetDolbyHeadphone',
            (['out'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method EnableDolbyDigitalLive')],
            HRESULT,
            'EnableDolbyDigitalLive',
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetDolbyDigitalLive')],
            HRESULT,
            'GetDolbyDigitalLive',
            (['out'], POINTER(LONG), 'pfEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method EnablePanorama')],
            HRESULT,
            'EnablePanorama',
            (['in'], LONG, 'fEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetPanorama')],
            HRESULT,
            'GetPanorama',
            (['out'], POINTER(LONG), 'pfEnable')
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
            (['out'], POINTER(ULONG), 'pdwCenterWidth')
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
            (['out'], POINTER(ULONG), 'pdwDimension')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetPrologicMode')],
            HRESULT,
            'SetPrologicMode',
            (['in'], ULONG, 'dwMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetPrologicMode')],
            HRESULT,
            'GetPrologicMode',
            (['out'], POINTER(ULONG), 'pdwMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetVirtualSpeakerMode')],
            HRESULT,
            'SetVirtualSpeakerMode',
            (['in'], ULONG, 'dwMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetVirtualSpeakerMode')],
            HRESULT,
            'GetVirtualSpeakerMode',
            (['out'], POINTER(ULONG), 'pdwMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetDolbyHeadphoneMode')],
            HRESULT,
            'SetDolbyHeadphoneMode',
            (['in'], ULONG, 'dwMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetDolbyHeadphoneMode')],
            HRESULT,
            'GetDolbyHeadphoneMode',
            (['out'], POINTER(ULONG), 'pdwMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetHeadphoneMode')],
            HRESULT,
            'GetHeadphoneMode',
            (['out'], POINTER(INT), 'pfHeadphoneMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsEnableDolbyPrologicChanged')],
            HRESULT,
            'IsEnableDolbyPrologicChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(LONG), 'pfEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsEnableDolbyVirtualSpeakerChanged')],
            HRESULT,
            'IsEnableDolbyVirtualSpeakerChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(LONG), 'pfEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsEnableDolbyHeadphoneChanged')],
            HRESULT,
            'IsEnableDolbyHeadphoneChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(LONG), 'pfEnabled')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsEnablePanoramaChanged')],
            HRESULT,
            'IsEnablePanoramaChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(LONG), 'pfEnabled')
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
            [comtypes.helpstring('method IsPrologicModeChanged')],
            HRESULT,
            'IsPrologicModeChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(ULONG), 'pdwMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsVirtualSpeakerModeChanged')],
            HRESULT,
            'IsVirtualSpeakerModeChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(ULONG), 'pdwMode')
        ),
        COMMETHOD(
            [comtypes.helpstring('method IsDolbyHeadphoneModeChanged')],
            HRESULT,
            'IsDolbyHeadphoneModeChanged',
            (['in'], PROPERTYKEY, 'key'),
            (['out'], POINTER(ULONG), 'pdwMode')
        ),
    ]


class DmsApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_DmsApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IDmsApi]

    @staticmethod
    def IDmsApi() -> IDmsApi:
        return comtypes.CoCreateInstance(
            CLSID_DmsApi,
            IDmsApi,
            comtypes.CLSCTX_ALL
        )
