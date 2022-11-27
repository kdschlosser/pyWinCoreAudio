
from ..data_types import *
from comtypes.automation import IDispatch
import comtypes
from .RtkApoApiLib import CLSID_RtkApoApiLib
from ..mmdeviceapi import IMMDevice


CLSID_RtApocapsApi = CLSID('{93EC11E9-889B-458E-8DEE-B750C8AAC6C2}')

IID_IRtApocapsApi = IID('{17D8D369-3B98-452F-AD26-8CEE32ACAA17}')
IID_IRtApocapsApiExt = IID('{3F28B387-BC4F-47E7-BC21-BED2C1DD3F88}')
IID_IRtApocapsApiExt2 = IID('{3AE98DD3-FFF2-405F-BFAF-D9BE5C788080}')
IID_IRtApocapsApiExt3 = IID('{0D8E5B7C-8DA2-4D29-BF26-017F61B2C711}')
IID_IRtApocapsApiExt4 = IID('{B4B135B1-8E1A-4E71-8687-E72B38524ABF}')
IID_IRtApocapsApiExt5 = IID('{45DE960C-F608-47C4-9D71-91CDAA057492}')
IID_IRtApocapsApiExt6 = IID('{922FE99B-E9B5-40E6-BC68-FA03547773A5}')
IID_IRtApocapsApiExt7 = IID('{6BE49DAF-0F34-40A3-9B8E-247FB658E062}')
IID_IRtApocapsApiExt8 = IID('{0BD9793C-4390-4827-8990-58D3A04DC3C1}')
IID_IRtApocapsApiExt9 = IID('{6B08C8FA-44E3-4546-8B76-EDCAC12C3A63}')
IID_IRtApocapsApiExt10 = IID('{A8980D44-CD9B-43B7-96EB-E222A02E1E45}')
IID_IRtApocapsApiExt11 = IID('{2EF1221A-6A67-4D28-8D7A-5A08344D4C08}')


class APOPOLICY(ctypes.Structure):
    _recordinfo_ = (
        CLSID_RtkApoApiLib, 1, 0, 0, '{ADACF436-CBD6-4A3D-8794-DA1D4AA6D8D3}'
    )
    _fields_ = [
        ('ulVersion', ULONG),
        ('Policy1', ULONG),
        ('Policy2', ULONG),
        ('Policy3', ULONG),
        ('Policy4', ULONG),
        ('Policy5', ULONG),
        ('Policy6', ULONG),
        ('Policy7', ULONG),
        ('Policy8', ULONG),
        ('Policy9', ULONG),
        ('Policy10', ULONG),
        ('Policy11', ULONG),
        ('Policy12', ULONG),
        ('Policy13', ULONG),
        ('Policy14', ULONG),
        ('Policy15', ULONG),
        ('Policy16', ULONG),
        ('Policy17', ULONG),
        ('Policy18', ULONG),
        ('Policy19', ULONG),
        ('Policy20', ULONG),
        ('Policy21', ULONG),
        ('Policy22', ULONG),
        ('Policy23', ULONG),
        ('Policy24', ULONG),
        ('Policy25', ULONG),
        ('Policy26', ULONG),
        ('Policy27', ULONG),
        ('Policy28', ULONG),
        ('Policy29', ULONG),
        ('Policy30', ULONG),
        ('Policy31', ULONG),
    ]


class APOFXCAPS3_COMPACT(ctypes.Structure):
    _recordinfo_ = (
        CLSID_RtkApoApiLib, 1, 0, 0, '{46EEB602-4428-4973-84C0-DC96FC393C6D}'
    )
    _fields_ = [
        ('ulVersion', ULONG),
        ('RtkFx', ULONG),
        ('RecFx', ULONG),
        ('MsFx', ULONG),
        ('OtherFx1', ULONG),
        ('OtherFx2', ULONG),
        ('OtherFx3', ULONG),
        ('OtherFx4', ULONG),
        ('OtherFx5', ULONG),
        ('ApoFx1', ULONG),
        ('ApoFx2', ULONG),
        ('ApoFx3', ULONG),
        ('ApoFx4', ULONG),
        ('ApoFx5', ULONG),
        ('ApoFx6', ULONG),
        ('ApoFx7', ULONG),
        ('ApoFx8', ULONG),
        ('ApoFx9', ULONG),
        ('ApoFx10', ULONG),
        ('ApoFx11', ULONG),
        ('ApoFx12', ULONG),
        ('ApoFx13', ULONG),
        ('ApoFx14', ULONG),
        ('ApoFx15', ULONG),
        ('ApoFx16', ULONG),
    ]


class IRtApocapsApi(IDispatch):
    _case_insensitive_ = True
    _iid_ = IID_IRtApocapsApi
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method Initialize')],
            HRESULT,
            'Initialize',
            (['in'], LPWSTR, 'pwszDeviceId'),
            (['in'], INT, 'nDataFlow')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetApoCaps')],
            HRESULT,
            'GetApoCaps',
            (['in'], INT, 'nGen'),
            (['out'], POINTER(UBYTE), 'pCaps'),
            (['in', 'out'], POINTER(UINT), 'cbSize')
        ),
        COMMETHOD(
            [comtypes.helpstring('method GetApoCapsByRole')],
            HRESULT,
            'GetApoCapsByRole',
            (['in'], INT, 'nGen'),
            (['in'], INT, 'nRoleBits'),
            (['out'], POINTER(UBYTE), 'pCaps'),
            (['in', 'out'], POINTER(UINT), 'cbSize')
        ),
    ]


class IRtApocapsApiExt(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IRtApocapsApiExt
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method SupportRawDefaultMode')],
            HRESULT,
            'SupportRawDefaultMode',
            (['out'], POINTER(INT), 'fSupport')
        ),
        COMMETHOD(
            [comtypes.helpstring('method EFXProxy')],
            HRESULT,
            'IsEFXProxy',
            (['out'], POINTER(INT), 'fProxy')
        ),
        COMMETHOD(
            [comtypes.helpstring('method MFXProxy')],
            HRESULT,
            'IsMFXProxy',
            (['out'], POINTER(INT), 'fProxy')
        ),
    ]


class IRtApocapsApiExt2(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IRtApocapsApiExt2
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetAudioPlatform')],
            HRESULT,
            'GetAudioPlatform',
            (['out'], POINTER(INT), 'pnPlatform')
        ),
    ]


class IRtApocapsApiExt3(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IRtApocapsApiExt3
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetLoadedApoGUID')],
            HRESULT,
            'GetLoadedApoGUID',
            (['in'], INT, 'nRoleBits'),
            (['out'], POINTER(POINTER(GUID)), 'ppAposIds'),
            (['out'], POINTER(UINT), 'pcApos')
        ),
    ]


class IRtApocapsApiExt4(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IRtApocapsApiExt4
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetApoGUID')],
            HRESULT,
            'GetApoGUID',
            (['in'], INT, 'nRoleBits'),
            (['out'], POINTER(POINTER(GUID)), 'ppAposIds'),
            (['out'], POINTER(POINTER(GUID)), 'ppAposInitIds'),
            (['out'], POINTER(POINTER(INT)), 'ppRoles'),
            (['out'], POINTER(UINT), 'pcApos')
        ),
    ]


class IRtApocapsApiExt5(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IRtApocapsApiExt5
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetEffectVersion')],
            HRESULT,
            'GetEffectVersion',
            (['in'], INT, 'nRoleBits'),
            (['in'], POINTER(GUID), 'pEffectGUID'),
            (['out'], POINTER(UINT), 'pVersionMS'),
            (['out'], POINTER(UINT), 'pVersionLS')
        ),
    ]


class IRtApocapsApiExt6(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IRtApocapsApiExt6
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method IsDCHUApoConfig')],
            HRESULT,
            'IsDCHUApoConfig',
            (['out'], POINTER(INT), 'fIsDCHU')
        )
    ]


class IRtApocapsApiExt7(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IRtApocapsApiExt7
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetAudioDllPath')],
            HRESULT,
            'GetAudioDllPath',
            (['out'], POINTER(WSTRING), 'ppPathName')
        ),
    ]


class IRtApocapsApiExt8(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IRtApocapsApiExt8
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetApoCapsCompact')],
            HRESULT,
            'GetApoCapsCompact',
            (['out'], POINTER(APOFXCAPS3_COMPACT), 'pCaps')
        ),
    ]


class IRtApocapsApiExt9(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IRtApocapsApiExt9
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetApoPolicy')],
            HRESULT,
            'GetApoPolicy',
            (['out'], POINTER(APOPOLICY), 'pApoPolicy')
        ),
    ]


class IRtApocapsApiExt10(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IRtApocapsApiExt10
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method GetApoPolicyEx')],
            HRESULT,
            'GetApoPolicyEx',
            (['in'], WSTRING, 'pwszHardwareId'),
            (['out'], POINTER(APOPOLICY), 'pApoPolicy')
        ),
        COMMETHOD(
            [comtypes.helpstring('method FindAllHwIDbyEpID')],
            HRESULT,
            'FindAllHwIDbyEpID',
            (['in'], WSTRING, 'pwszEndpointId'),
            (['out'], POINTER(WSTRING), 'ppwszHardwareId'),
            (['out'], POINTER(INT), 'pnWCharCount')
        ),
    ]


class IRtApocapsApiExt11(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IRtApocapsApiExt11
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    _methods_ = [
        COMMETHOD(
            [comtypes.helpstring('method FindAllHwIDbySoftIODev')],
            HRESULT,
            'FindAllHwIDbySoftIODev',
            (['in'], POINTER(IMMDevice), 'pSoftwareIoDevice'),
            (['out'], POINTER(WSTRING), 'ppwszHardwareId'),
            (['out'], POINTER(INT), 'pnWCharCount')
        ),
    ]


class RtApocapsApi(comtypes.CoClass):
    _reg_clsid_ = CLSID_RtApocapsApi
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = (CLSID_RtkApoApiLib, 1, 0)
    _com_interfaces_ = [
        IRtApocapsApi,
        IRtApocapsApiExt,
        IRtApocapsApiExt2,
        IRtApocapsApiExt3,
        IRtApocapsApiExt4,
        IRtApocapsApiExt5,
        IRtApocapsApiExt6,
        IRtApocapsApiExt7,
        IRtApocapsApiExt8,
        IRtApocapsApiExt9,
        IRtApocapsApiExt10,
        IRtApocapsApiExt11
    ]

    @staticmethod
    def IRtApocapsApi() -> IRtApocapsApi:
        return comtypes.CoCreateInstance(
            CLSID_RtApocapsApi,
            IRtApocapsApi,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IRtApocapsApiExt() -> IRtApocapsApiExt:
        return comtypes.CoCreateInstance(
            CLSID_RtApocapsApi,
            IRtApocapsApiExt,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IRtApocapsApiExt2() -> IRtApocapsApiExt2:
        return comtypes.CoCreateInstance(
            CLSID_RtApocapsApi,
            IRtApocapsApiExt2,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IRtApocapsApiExt3() -> IRtApocapsApiExt3:
        return comtypes.CoCreateInstance(
            CLSID_RtApocapsApi,
            IRtApocapsApiExt3,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IRtApocapsApiExt4() -> IRtApocapsApiExt4:
        return comtypes.CoCreateInstance(
            CLSID_RtApocapsApi,
            IRtApocapsApiExt4,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IRtApocapsApiExt5() -> IRtApocapsApiExt5:
        return comtypes.CoCreateInstance(
            CLSID_RtApocapsApi,
            IRtApocapsApiExt5,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IRtApocapsApiExt6() -> IRtApocapsApiExt6:
        return comtypes.CoCreateInstance(
            CLSID_RtApocapsApi,
            IRtApocapsApiExt6,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IRtApocapsApiExt7() -> IRtApocapsApiExt7:
        return comtypes.CoCreateInstance(
            CLSID_RtApocapsApi,
            IRtApocapsApiExt7,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IRtApocapsApiExt8() -> IRtApocapsApiExt8:
        return comtypes.CoCreateInstance(
            CLSID_RtApocapsApi,
            IRtApocapsApiExt8,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IRtApocapsApiExt9() -> IRtApocapsApiExt9:
        return comtypes.CoCreateInstance(
            CLSID_RtApocapsApi,
            IRtApocapsApiExt9,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IRtApocapsApiExt10() -> IRtApocapsApiExt10:
        return comtypes.CoCreateInstance(
            CLSID_RtApocapsApi,
            IRtApocapsApiExt10,
            comtypes.CLSCTX_ALL
        )

    @staticmethod
    def IRtApocapsApiExt11() -> IRtApocapsApiExt11:
        return comtypes.CoCreateInstance(
            CLSID_RtApocapsApi,
            IRtApocapsApiExt11,
            comtypes.CLSCTX_ALL
        )

