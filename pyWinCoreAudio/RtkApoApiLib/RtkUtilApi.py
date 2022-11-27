from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib

CLSID_RtkUtilApi = CLSID('{E4D7C1C8-9459-45F8-A541-CE54DE835DD5}')
IID_IHardwareIdApi = IID('{93BFDA87-6364-4E43-B7BF-2DFD0FA6A902}')
IID_IRtkApoApiVersion = IID('{B477AF0F-2CA0-450B-B5E3-3C56E14DB83B}')

IID_ISupportedProgsApi = IID('{FA68D0EA-DEE7-468B-8BB3-EAA4F023E391}')


class IHardwareIdApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IHardwareIdApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.dispid(1)],
            HRESULT,
            'GetPciId',
            (['out'], POINTER(USHORT), 'pwPciSVID'),
            (['out'], POINTER(USHORT), 'pwPciSSID')
        ),
        COMMETHOD(
            [comtypes.dispid(2)],
            HRESULT,
            'GetCodecId',
            (['out'], POINTER(USHORT), 'pwCodecSVID'),
            (['out'], POINTER(USHORT), 'pwCodecSSID')
        ),
    ]


class ISupportedProgsApi(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_ISupportedProgsApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.dispid(1)],
            HRESULT,
            'CheckProgramSupport',
            (['in'], INT, 'nProgId'),
            (['out'], POINTER(INT), 'fSupported')
        ),
    ]


class IRtkApoApiVersion(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IRtkApoApiVersion
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.dispid(1)],
            HRESULT,
            'GetVersion',
            (['out', 'retval'], POINTER(INT), 'pnVersion')
        ),
    ]


class RtkUtilApi(comtypes.CoClass):
    """RtkUtilApi Class"""
    _reg_clsid_ = CLSID_RtkUtilApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IHardwareIdApi, ISupportedProgsApi, IRtkApoApiVersion]

    @staticmethod
    def IHardwareIdApi() -> IHardwareIdApi:
        return comtypes.CoCreateInstance(
            CLSID_RtkUtilApi,
            IHardwareIdApi,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def ISupportedProgsApi() -> ISupportedProgsApi:
        return comtypes.CoCreateInstance(
            CLSID_RtkUtilApi,
            ISupportedProgsApi,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IRtkApoApiVersion() -> IRtkApoApiVersion:
        return comtypes.CoCreateInstance(
            CLSID_RtkUtilApi,
            IRtkApoApiVersion,
            comtypes.CLSCTX_ALL
        )
