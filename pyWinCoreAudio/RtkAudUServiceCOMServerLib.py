from .data_types import *
import comtypes
from comtypes.automation import IDispatch


class _IRtkAudUServiceCOMObjectEvents(IDispatch):
    _case_insensitive_ = True
    '_IRtkAudUServiceCOMObjectEvents Interface'
    _iid_ = GUID('{EA1C918E-7C98-4434-B18A-4EE83C866FF8}')
    _idlflags_ = []
    _methods_ = []
    
    
_IRtkAudUServiceCOMObjectEvents._disp_methods_ = [
    comtypes.DISPMETHOD(
        [comtypes.dispid(1), comtypes.helpstring('method OnRtkAudioServiceEvent')], 
        HRESULT,
        'OnRtkAudioServiceEvent',
        (['in'], ULONG, 'nCommand')
    ),
]


class Library(object):
    """RtkAudUServiceCOMServer 1.0 Type Library"""
    name = 'RtkAudUServiceCOMServerLib'
    _reg_typelib_ = ('{71092E27-C34B-47D4-A961-8AD778F68191}', 1, 0)


class RtkAudUServiceCOMObject(comtypes.CoClass):
    """RtkAudUServiceCOMObject Class"""
    _reg_clsid_ = GUID('{615AC66B-72C3-4DEB-8F22-19B372142787}')
    _idlflags_ = []
    # _typelib_path_ = typelib_path
    _reg_typelib_ = ('{71092E27-C34B-47D4-A961-8AD778F68191}', 1, 0)
    
    
class IRtkAudUServiceCOMObject(IDispatch):
    """IRtkAudUServiceCOMObject Interface"""
    _case_insensitive_ = True
    _iid_ = GUID('{5FF6484C-1BC1-49E6-B61F-A207D656CBDC}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    
    
RtkAudUServiceCOMObject._com_interfaces_ = [IRtkAudUServiceCOMObject]
RtkAudUServiceCOMObject._outgoing_interfaces_ = [_IRtkAudUServiceCOMObjectEvents]


class IRtkAudUServiceCOMNotificationClient(comtypes.IUnknown):
    """IRtkAudUServiceCOMNotificationClient Interface"""
    _case_insensitive_ = True
    _iid_ = GUID('{572D0BCA-0207-4162-A74D-848870852D82}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
    
    
IRtkAudUServiceCOMObject._methods_ = [
    comtypes.COMMETHOD(
        [comtypes.dispid(1), comtypes.helpstring('method AllowAccess')], 
        HRESULT, 
        'AllowAccess',
        (['in'], INT, 'nAccessCode')
    ),
    comtypes.COMMETHOD(
        [comtypes.dispid(2), comtypes.helpstring('method ServiceCommandWithIndex')], 
        HRESULT, 
        'ServiceCommandWithIndex',
        (['in'], WSTRING, 'szPath'),
        (['in'], ULONG, 'nCommand'),
        (['in'], INT, 'nIndex'),
        (['in', 'out'], POINTER(UBYTE), 'lpInBuffer'),
        (['in'], ULONG, 'nInBufferSize'),
        (['in', 'out'], POINTER(UBYTE), 'lpOutBuffer'),
        (['in'], ULONG, 'nOutBufferSize'),
        (['out'], POINTER(ULONG), 'ulBytesReturned')
    ),
    comtypes.COMMETHOD(
        [comtypes.dispid(3), comtypes.helpstring('method ServiceCommandIntWithIndex')], 
        HRESULT, 
        'ServiceCommandIntWithIndex',
        (['in'], WSTRING, 'szPath'),
        (['in'], ULONG, 'nCommand'),
        (['in'], INT, 'nIndex'),
        (['in'], INT, 'nInValue'),
        (['out'], POINTER(INT), 'nOutValue')
    ),
    comtypes.COMMETHOD(
        [comtypes.dispid(4), comtypes.helpstring('method ServiceCommandFloatWithIndex')], 
        HRESULT, 
        'ServiceCommandFloatWithIndex',
        (['in'], WSTRING, 'szPath'),
        (['in'], ULONG, 'nCommand'),
        (['in'], INT, 'nIndex'),
        (['in'], FLOAT, 'fInValue'),
        (['out'], POINTER(FLOAT), 'fOutValue')
    ),
    comtypes.COMMETHOD(
        [comtypes.dispid(5), comtypes.helpstring('method RegisterServiceEvent')], 
        HRESULT, 
        'RegisterServiceEvent',
        (['in'], WSTRING, 'szPath'),
        (['in'], ULONG, 'nCommand')
    ),
    comtypes.COMMETHOD(
        [comtypes.dispid(6), comtypes.helpstring('method UnregisterServiceEvent')], 
        HRESULT, 
        'UnregisterServiceEvent',
        (['in'], WSTRING, 'szPath'),
        (['in'], ULONG, 'nCommand')
    ),
    comtypes.COMMETHOD(
        [comtypes.dispid(7), comtypes.helpstring('method RegisterRtkAudUServiceCOMNotificationCallback')], 
        HRESULT, 
        'RegisterRtkAudUServiceCOMNotificationCallback',
        (['in'], POINTER(IRtkAudUServiceCOMNotificationClient), 'pClient')
    ),
    comtypes.COMMETHOD(
        [comtypes.dispid(8), comtypes.helpstring('method UnregisterRtkAudUServiceCOMNotificationCallback')], 
        HRESULT, 
        'UnregisterRtkAudUServiceCOMNotificationCallback',
        (['in'], POINTER(IRtkAudUServiceCOMNotificationClient), 'pClient')
    ),
]
################################################################
## code template for IRtkAudUServiceCOMObject implementation
##class IRtkAudUServiceCOMObject_Impl(object):
##    def AllowAccess(self, nAccessCode):
##        'method AllowAccess'
##        #return 
##
##    def ServiceCommandWithIndex(self, szPath, nCommand, nIndex, nInBufferSize, nOutBufferSize):
##        'method ServiceCommandWithIndex'
##        #return lpInBuffer, lpOutBuffer, ulBytesReturned
##
##    def ServiceCommandIntWithIndex(self, szPath, nCommand, nIndex, nInValue):
##        'method ServiceCommandIntWithIndex'
##        #return nOutValue
##
##    def ServiceCommandFloatWithIndex(self, szPath, nCommand, nIndex, fInValue):
##        'method ServiceCommandFloatWithIndex'
##        #return fOutValue
##
##    def RegisterServiceEvent(self, szPath, nCommand):
##        'method RegisterServiceEvent'
##        #return 
##
##    def UnregisterServiceEvent(self, szPath, nCommand):
##        'method UnregisterServiceEvent'
##        #return 
##
##    def RegisterRtkAudUServiceCOMNotificationCallback(self, pClient):
##        'method RegisterRtkAudUServiceCOMNotificationCallback'
##        #return 
##
##    def UnregisterRtkAudUServiceCOMNotificationCallback(self, pClient):
##        'method UnregisterRtkAudUServiceCOMNotificationCallback'
##        #return 
##

IRtkAudUServiceCOMNotificationClient._methods_ = [
    comtypes.COMMETHOD(
        [comtypes.dispid(1), comtypes.helpstring('method OnAudioDeviceEvent')], 
        HRESULT, 
        'OnAudioDeviceEvent',
        (['in'], WSTRING, 'szPath'),
        (['in'], ULONG, 'nCommand')
    ),
    comtypes.COMMETHOD(
        [comtypes.dispid(2), comtypes.helpstring('method OnServiceEvent')], 
        HRESULT, 
        'OnServiceEvent',
        (['in'], ULONG, 'nCommand')
    ),
]
################################################################
## code template for IRtkAudUServiceCOMNotificationClient implementation
##class IRtkAudUServiceCOMNotificationClient_Impl(object):
##    def OnAudioDeviceEvent(self, szPath, nCommand):
##        'method OnAudioDeviceEvent'
##        #return 
##
##    def OnServiceEvent(self, nCommand):
##        'method OnServiceEvent'
##        #return 
##

__all__ = [ 'RtkAudUServiceCOMObject', 'IRtkAudUServiceCOMObject',
           'IRtkAudUServiceCOMNotificationClient',
           '_IRtkAudUServiceCOMObjectEvents']
