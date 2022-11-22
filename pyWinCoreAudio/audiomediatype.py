
from .data_types import *
import ctypes
import comtypes
from .audioclient import PWAVEFORMATEX


class _UNCOMPRESSEDAUDIOFORMAT(ctypes.Structure):
    _fields_ = [
        ('guidFormatType', GUID),
        ('dwSamplesPerFrame', DWORD),
        ('dwBytesPerSampleContainer', DWORD),
        ('dwValidBitsPerSample', DWORD),
        ('fFramesPerSecond', FLOAT),
        ('dwChannelMask', DWORD)
    ]


UNCOMPRESSEDAUDIOFORMAT = _UNCOMPRESSEDAUDIOFORMAT

IID_IAudioMediaType = IID(
    '{4E997F73-B71F-4798-873B-ED7DFCF15B4D}'
)


class IAudioMediaType(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioMediaType

    
IAudioMediaType._methods_ = (
    COMMETHOD(
        [],
        HRESULT,
        'IsCompressedFormat',
        (['out'], POINTER(BOOL), 'pfCompressed')
    ),
    COMMETHOD(
        [],
        HRESULT,
        'IsEqual',
        (['in'], POINTER(IAudioMediaType), 'pIAudioType'),
        (['out'], POINTER(DWORD), 'pdwFlags')
    ),
    COMMETHOD(
        [],
        PWAVEFORMATEX,
        'GetAudioFormat',
    ),
    COMMETHOD(
        [],
        HRESULT,
        'GetUncompressedAudioFormat',
        (['out'], POINTER(UNCOMPRESSEDAUDIOFORMAT), 'pUncompressedAudioFormat')
    ),
)

# CreateAudioMediaType
# 
# STDAPI CreateAudioMediaType(
#     const WAVEFORMATEX* pAudioFormat,
#     UINT32 cbAudioFormatSize,
#     IAudioMediaType** ppIAudioMediaType
#     );

# CreateAudioMediaTypeFromUncompressedAudioFormat
# 
# STDAPI CreateAudioMediaTypeFromUncompressedAudioFormat(
#     const UNCOMPRESSEDAUDIOFORMAT* pUncompressedAudioFormat,
#     IAudioMediaType** ppIAudioMediaType
#     );
#     

AUDIOMEDIATYPE_EQUAL_FORMAT_TYPES = 0x00000002
AUDIOMEDIATYPE_EQUAL_FORMAT_DATA = 0x00000004
AUDIOMEDIATYPE_EQUAL_FORMAT_USER_DATA = 0x00000008
