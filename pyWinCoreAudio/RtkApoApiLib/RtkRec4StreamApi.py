from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib
from ..mmreg import WAVEFORMATEX


CLSID_RtkRec4StreamApi = CLSID('{3833A070-8C55-4465-9B8F-BDBE0C6498B8}')
IID_IRtkRec4StreamApi = IID('{9DC44DE9-7749-4558-BADF-4C90AFE88089}')


class IRtkRec4StreamApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IRtkRec4StreamApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method Initialize')],
            HRESULT,
            'Initialize'
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetMixFormat')],
            HRESULT,
            'GetMixFormat',
            (['in'], INT, 'nStreamIndex'),
            (['in', 'out'], POINTER(POINTER(WAVEFORMATEX)), 'ppDeviceFormat')
        ),
        COMMETHOD(
            [comtypes.helpstring('method SetEventHandle')],
            HRESULT,
            'SetEventHandle',
            (['in'], INT, 'nStreamIndex'),
            (['in'], VOID, 'eventHandle')
        ),
        COMMETHOD(
            [comtypes.helpstring('method Start')],
            HRESULT,
            'Start'
        ),
        COMMETHOD(
            [comtypes.helpstring('method Stop')],
            HRESULT,
            'Stop'
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetNextPacketSize')],
            HRESULT,
            'GetNextPacketSize',
            (['in'], INT, 'nStreamIndex'),
            (['in', 'out'], POINTER(UINT), 'pNumFramesInNextPacket')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetBuffer')],
            HRESULT,
            'GetBuffer',
            (['in'], INT, 'nStreamIndex'),
            (['in', 'out'], POINTER(POINTER(UBYTE)), 'ppData'),
            (['in', 'out'], POINTER(UINT), 'pNumFramesToRead'),
            (['in', 'out'], POINTER(ULONG), 'pdwFlags'),
            (['in', 'out'], POINTER(ULONGLONG), 'pu64DevicePosition'),
            (['in', 'out'], POINTER(ULONGLONG), 'pu64QPCPosition')
        ),
        COMMETHOD(
            [comtypes.helpstring('method ReleaseBuffer')],
            HRESULT,
            'ReleaseBuffer',
            (['in'], INT, 'nStreamIndex'),
            (['in'], UINT, 'NumFramesRead')
        ),
    ]


class RtkRec4StreamApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_RtkRec4StreamApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IRtkRec4StreamApi]

    @staticmethod
    def IRtkRec4StreamApi() -> IRtkRec4StreamApi:
        return comtypes.CoCreateInstance(
            CLSID_RtkRec4StreamApi,
            IRtkRec4StreamApi,
            comtypes.CLSCTX_ALL
        )
