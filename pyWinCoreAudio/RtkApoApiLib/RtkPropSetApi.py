from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_RtkPropSetApi = CLSID('{ACF17F45-928A-42A7-8591-B64CBA5EED63}')
IID_IPropSet = IID('{2B1098F3-7079-4F45-92E1-7548A7406345}')
IID_IPropSet20065 = IID('{58F01532-7504-4D49-8428-38136895D1D6}')
IID_IPropSet20082 = IID('{8BFB6CBD-EDC6-4CBA-9B56-F59DF6E72291}')
IID_IPropSet20098 = IID('{A19A0112-BF84-41C4-9891-541C0F2F35EC}')
IID_IPropSet20136 = IID('{608ED38D-7ABF-4071-8BDF-5312DC214426}')
IID_IPropSet20140 = IID('{6CEEABA2-BA60-4302-9CE6-8C0FA322D9AE}')


class IPropSet(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IPropSet
    _idlflags_ = ['dual', 'oleautomation']
    _methods_ = []


class IPropSet20065(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropSet20065
    _idlflags_ = ['dual', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetDeviceCount')],
            HRESULT,
            'GetDeviceCount',
            (['in', 'out'], POINTER(ULONG), 'ulCount')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetDeviceName')],
            HRESULT,
            'GetDeviceName',
            (['in'], ULONG, 'ulIndex'),
            (['out'], POINTER(USHORT), 'cDeviceName')
        ),
        COMMETHOD(
            [comtypes.helpstring('method VistaGetEqEnable')],
            HRESULT,
            'VistaGetEqEnable',
            (['in'], ULONG, 'ulDeviceIndex'),
            (['out'], POINTER(INT), 'bEqEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method VistaSetEqEnable')],
            HRESULT,
            'VistaSetEqEnable',
            (['in'], ULONG, 'ulDeviceIndex'),
            (['in'], INT, 'bEqEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method VistaGetEqLevel')],
            HRESULT,
            'VistaGetEqLevel',
            (['in'], ULONG, 'ulDeviceIndex'),
            (['in'], ULONG, 'ulEqIndex'),
            (['out'], POINTER(INT), 'lEqLevel')
        ),
        COMMETHOD(
            [comtypes.helpstring('method VistaSetEqLevel')],
            HRESULT,
            'VistaSetEqLevel',
            (['in'], ULONG, 'ulDeviceIndex'),
            (['in'], ULONG, 'ulEqIndex'),
            (['in'], INT, 'lEqLevel')
        ),
    ]


class IPropSet20082(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropSet20082
    _idlflags_ = ['dual', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method VistaGetDeviceInfo')],
            HRESULT,
            'VistaGetDeviceInfo',
            (['in'], ULONG, 'ulDeviceIndex'),
            (['out'], POINTER(ULONG), 'DeviceFormFactor'),
            (['out'], POINTER(INT), 'bIsPlayback'),
            (['out'], POINTER(INT), 'bIsDefDevice'),
            (['out'], POINTER(USHORT), 'wcDeviceID')
        ),
    ]


class IPropSet20098(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropSet20098
    _idlflags_ = ['dual', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method VistaGetForteMicBeamForming')],
            HRESULT,
            'VistaGetForteMicBeamForming',
            (['in'], ULONG, 'ulDeviceIndex'),
            (['out'], POINTER(INT), 'bEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method VistaSetForteMicBeamForming')],
            HRESULT,
            'VistaSetForteMicBeamForming',
            (['in'], ULONG, 'ulDeviceIndex'),
            (['in'], INT, 'bEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method VistaGetForteMicAEC')],
            HRESULT,
            'VistaGetForteMicAEC',
            (['in'], ULONG, 'ulDeviceIndex'),
            (['out'], POINTER(INT), 'bEnable')
        ),
        COMMETHOD(
            [comtypes.helpstring('method VistaSetForteMicAEC')],
            HRESULT,
            'VistaSetForteMicAEC',
            (['in'], ULONG, 'ulDeviceIndex'),
            (['in'], INT, 'bEnable')
        ),
    ]


class IPropSet20136(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropSet20136
    _idlflags_ = ['dual', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetHWVolumeStatus')],
            HRESULT,
            'GetHWVolumeStatus',
            (['in', 'out'], POINTER(INT), 'bStatus')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetHWVolumeStatus')],
            HRESULT,
            'SetHWVolumeStatus',
            (['in'], INT, 'bStatus')
        ),
    ]


class IPropSet20140(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropSet20140
    _idlflags_ = ['dual', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetForteFFPVolume')],
            HRESULT,
            'GetForteFFPVolume',
            (['out'], POINTER(INT), 'Vol')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetForteFFPVolume')],
            HRESULT,
            'SetForteFFPVolume',
            (['in'], INT, 'Vol')
        ),
    ]


class RtkPropSetApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_RtkPropSetApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [
        IPropSet,
        IPropSet20065,
        IPropSet20082,
        IPropSet20098,
        IPropSet20136,
        IPropSet20140
    ]

    @staticmethod
    def IPropSet() -> IPropSet:
        return comtypes.CoCreateInstance(
            CLSID_RtkPropSetApi,
            IPropSet,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IPropSet20065() -> IPropSet20065:
        return comtypes.CoCreateInstance(
            CLSID_RtkPropSetApi,
            IPropSet20065,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IPropSet20082() -> IPropSet20082:
        return comtypes.CoCreateInstance(
            CLSID_RtkPropSetApi,
            IPropSet20082,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IPropSet20098() -> IPropSet20098:
        return comtypes.CoCreateInstance(
            CLSID_RtkPropSetApi,
            IPropSet20098,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IPropSet20136() -> IPropSet20136:
        return comtypes.CoCreateInstance(
            CLSID_RtkPropSetApi,
            IPropSet20136,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IPropSet20140() -> IPropSet20140:
        return comtypes.CoCreateInstance(
            CLSID_RtkPropSetApi,
            IPropSet20140,
            comtypes.CLSCTX_ALL
        )
