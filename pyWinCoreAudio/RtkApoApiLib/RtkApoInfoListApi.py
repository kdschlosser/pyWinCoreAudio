from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib
from comtypes.automation import _midlSAFEARRAY  # NOQA
from ..mmreg import WAVEFORMATEXTENSIBLE


CLSID_RtkApoInfoListApi = CLSID('{AA9642C4-CE9B-4564-A57E-AC5B8F3E7D77}')
IID_IRtkApoInfoListApi = IID('{22D842FF-478F-4FA2-A66F-E34C417CAABB}')


class APO_INSTANCE(ctypes.Structure):
    _recordinfo_ = (
        CLSID_RtkApoApiLib, 1, 0, 0, '{B83FB95B-6043-49F1-BA4D-505A12A75FFF}'
    )
    _fields_ = [
        ('nInstanceIndex', INT),
        ('nDataFlow', INT),
        ('EndpointId', GUID),
        ('ClassId', GUID),
        ('ProcessingMode', GUID),
    ]


class APO_INFO(ctypes.Structure):
    _recordinfo_ = (
        CLSID_RtkApoApiLib, 1, 0, 0, '{AAD67C92-B9DD-4510-9352-D059D1C7F604}'
    )
    _fields_ = [
        ('ClassId', GUID),
        ('wfxInput', WAVEFORMATEXTENSIBLE),
        ('wfxOutput', WAVEFORMATEXTENSIBLE),
        ('nodeType', INT),
        ('parameters', INT * 256),
    ]


class APO_EDGE(ctypes.Structure):
    _recordinfo_ = (
        CLSID_RtkApoApiLib, 1, 0, 0, '{264F145B-A7B7-408F-9000-807688BDE4D5}'
    )
    _fields_ = [
        ('nFrom', INT),
        ('nTo', INT),
    ]


class IRtkApoInfoListApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IRtkApoInfoListApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetApoInstances')],
            HRESULT,
            'GetApoInstances',
            (['out'], POINTER(_midlSAFEARRAY(APO_INSTANCE)), 'instArr')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetApoInfoList')],
            HRESULT,
            'GetApoInfoList',
            (['in'], POINTER(APO_INSTANCE), 'pApoIns'),
            (['out'], POINTER(_midlSAFEARRAY(APO_INFO)), 'effectsArr'),
            (['out'], POINTER(_midlSAFEARRAY(APO_EDGE)), 'edgesArr')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetEffectName')],
            HRESULT,
            'GetEffectName',
            (['in'], GUID, 'effectGuid'),
            (['out'], POINTER(BSTR), 'psName')
        ),
    ]


class RtkApoInfoListApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_RtkApoInfoListApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [IRtkApoInfoListApi]

    @staticmethod
    def IRtkApoInfoListApi() -> IRtkApoInfoListApi:
        return comtypes.CoCreateInstance(
            CLSID_RtkApoInfoListApi,
            IRtkApoInfoListApi,
            comtypes.CLSCTX_ALL
        )
