from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib


CLSID_RtkApoNotification = CLSID('{0E249A19-94DF-418A-9191-41D1FA0F3C2E}')
IID_IRtkApoNotificationClient = IID('{A946DB9E-5685-4938-81D5-291E6A31D8D5}')
IID_IRtkApoNotification = IID('{C4B85D3D-9573-4864-96C5-FA7ABF303F1E}')


class IRtkApoNotificationClient(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IRtkApoNotificationClient
    _idlflags_ = ['nonextensible']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method OnPropertyValueChanged')],
            HRESULT,
            'OnPropertyValueChanged',
            (['in'], ULONG, 'dwNotify')
        ),
    ]


class IRtkApoNotification(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IRtkApoNotification
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method Initialize')],
            HRESULT,
            'Initialize',
            (['in'], LPWSTR, 'pwszDeviceId')
        ),
        COMMETHOD(
            [comtypes.helpstring('method RegisterNotificationCallback')],
            HRESULT,
            'RegisterNotificationCallback',
            (['in'], POINTER(IRtkApoNotificationClient), 'pNotify')
        ),
        COMMETHOD(
            [comtypes.helpstring('method UnregisterNotificationCallback')],
            HRESULT,
            'UnregisterNotificationCallback',
            (['in'], POINTER(IRtkApoNotificationClient), 'pNotify')
        ),
    ]


class RtkApoNotification(comtypes.CoClass):
    _reg_clsid_ = CLSID_RtkApoNotification
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IRtkApoNotification]

    @staticmethod
    def IRtkApoNotification() -> IRtkApoNotification:
        return comtypes.CoCreateInstance(
            CLSID_RtkApoNotification,
            IRtkApoNotification,
            comtypes.CLSCTX_ALL
        )
