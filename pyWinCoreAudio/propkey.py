
from .guiddef import PROPERTYKEY as _PROPERTYKEY, GUID


class PK(_PROPERTYKEY):
    def __str__(self):
        for name, pkey in globals().items():
            if not name.startswith('PKEY_'):
                continue

            if (
                str(pkey.fmtid) == str(self.fmtid) and
                pkey.pid == self.pid
            ):
                return name

        return _PROPERTYKEY.__str__(self)


def DEFINE_PROPERTYKEY(l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8, key):
    return PK(GUID(l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8), key)


# Name:     System.Audio.ChannelCount -- PKEY_Audio_ChannelCount
# Type:     UInt32 -- VT_UI4
# FormatID: (FMTID_AudioSummaryInformation)
# {64440490-4C8B-11D1-8B70-080036B11A03}, 7 (PIDASI_CHANNEL_COUNT)
# Indicates the channel count for the audio file.  Values: 1 (mono), 2 (stereo).
PKEY_Audio_ChannelCount = DEFINE_PROPERTYKEY(
    0x64440490,
    0x4C8B,
    0x11D1,
    0x8B,
    0x70,
    0x08,
    0x00,
    0x36,
    0xB1,
    0x1A,
    0x03,
    7
)

# Possible discrete values for PKEY_Audio_ChannelCount are:
AUDIO_CHANNELCOUNT_MONO = 1
AUDIO_CHANNELCOUNT_STEREO = 2

# Name:     System.Audio.Compression -- PKEY_Audio_Compression
# Type:     String -- VT_LPWSTR  (For variants: VT_BSTR)
# FormatID: (FMTID_AudioSummaryInformation)
# {64440490-4C8B-11D1-8B70-080036B11A03}, 10 (PIDASI_COMPRESSION)
PKEY_Audio_Compression = DEFINE_PROPERTYKEY(
    0x64440490,
    0x4C8B,
    0x11D1,
    0x8B,
    0x70,
    0x08,
    0x00,
    0x36,
    0xB1,
    0x1A,
    0x03,
    10
)

# Name:     System.Audio.EncodingBitrate -- PKEY_Audio_EncodingBitrate
# Type:     UInt32 -- VT_UI4
# FormatID: (FMTID_AudioSummaryInformation)
# {64440490-4C8B-11D1-8B70-080036B11A03}, 4 (PIDASI_AVG_DATA_RATE)
# Indicates the average data rate in Hz for the audio file in "bits per second".
PKEY_Audio_EncodingBitrate = DEFINE_PROPERTYKEY(
    0x64440490,
    0x4C8B,
    0x11D1,
    0x8B,
    0x70,
    0x08,
    0x00,
    0x36,
    0xB1,
    0x1A,
    0x03,
    4
)

# Name:     System.Audio.Format -- PKEY_Audio_Format
# Type:     String -- VT_LPWSTR  (For variants: VT_BSTR)
# Legacy code may treat this as VT_BSTR.
# FormatID: (FMTID_AudioSummaryInformation)
# {64440490-4C8B-11D1-8B70-080036B11A03}, 2 (PIDASI_FORMAT)
# Indicates the format of the audio file.
PKEY_Audio_Format = DEFINE_PROPERTYKEY(
    0x64440490,
    0x4C8B,
    0x11D1,
    0x8B,
    0x70,
    0x08,
    0x00,
    0x36,
    0xB1,
    0x1A,
    0x03,
    2
)


# Name:     System.Audio.IsVariableBitRate -- PKEY_Audio_IsVariableBitRate
# Type:     Boolean -- VT_BOOL
# FormatID: {E6822FEE-8C17-4D62-823C-8E9CFCBD1D5C}, 100
PKEY_Audio_IsVariableBitRate = DEFINE_PROPERTYKEY(
    0xE6822FEE,
    0x8C17,
    0x4D62,
    0x82,
    0x3C,
    0x8E,
    0x9C,
    0xFC,
    0xBD,
    0x1D,
    0x5C,
    100
)

# Name:     System.Audio.PeakValue -- PKEY_Audio_PeakValue
# Type:     UInt32 -- VT_UI4
# FormatID: {2579E5D0-1116-4084-BD9A-9B4F7CB4DF5E}, 100
PKEY_Audio_PeakValue = DEFINE_PROPERTYKEY(
    0x2579E5D0,
    0x1116,
    0x4084,
    0xBD,
    0x9A,
    0x9B,
    0x4F,
    0x7C,
    0xB4,
    0xDF,
    0x5E,
    100
)

# Name:     System.Audio.SampleRate -- PKEY_Audio_SampleRate
# Type:     UInt32 -- VT_UI4
# FormatID: (FMTID_AudioSummaryInformation)
# {64440490-4C8B-11D1-8B70-080036B11A03}, 5 (PIDASI_SAMPLE_RATE)
# Indicates the audio sample rate for the audio file in "samples per second".
PKEY_Audio_SampleRate = DEFINE_PROPERTYKEY(
    0x64440490,
    0x4C8B,
    0x11D1,
    0x8B,
    0x70,
    0x08,
    0x00,
    0x36,
    0xB1,
    0x1A,
    0x03,
    5
)

# Name:     System.Audio.SampleSize -- PKEY_Audio_SampleSize
# Type:     UInt32 -- VT_UI4
# FormatID: (FMTID_AudioSummaryInformation)
# {64440490-4C8B-11D1-8B70-080036B11A03}, 6 (PIDASI_SAMPLE_SIZE)
# Indicates the audio sample size for the audio file in "bits per sample".
PKEY_Audio_SampleSize = DEFINE_PROPERTYKEY(
    0x64440490,
    0x4C8B,
    0x11D1,
    0x8B,
    0x70,
    0x08,
    0x00,
    0x36,
    0xB1,
    0x1A,
    0x03,
    6
)

# Name:     System.Audio.StreamName -- PKEY_Audio_StreamName
# Type:     String -- VT_LPWSTR  (For variants: VT_BSTR)
# FormatID: (FMTID_AudioSummaryInformation)
# {64440490-4C8B-11D1-8B70-080036B11A03}, 9 (PIDASI_STREAM_NAME)
PKEY_Audio_StreamName = DEFINE_PROPERTYKEY(
    0x64440490,
    0x4C8B,
    0x11D1,
    0x8B,
    0x70,
    0x08,
    0x00,
    0x36,
    0xB1,
    0x1A,
    0x03,
    9
)

# Name:     System.Audio.StreamNumber -- PKEY_Audio_StreamNumber
# Type:     UInt16 -- VT_UI2
# FormatID: (FMTID_AudioSummaryInformation)
# {64440490-4C8B-11D1-8B70-080036B11A03}, 8 (PIDASI_STREAM_NUMBER)
PKEY_Audio_StreamNumber = DEFINE_PROPERTYKEY(
    0x64440490,
    0x4C8B,
    0x11D1,
    0x8B,
    0x70,
    0x08,
    0x00,
    0x36,
    0xB1,
    0x1A,
    0x03,
    8
)

# Name:     System.Devices.AudioDevice.Microphone.IsFarField --
# PKEY_Devices_AudioDevice_Microphone_IsFarField
# Type:     Boolean -- VT_BOOL
# FormatID: {8943B373-388C-4395-B557-BC6DBAFFAFDB}, 6
# Far field capability of the microphone.
# If VARIANT_TRUE the microphone element will detect far field sound.
PKEY_Devices_AudioDevice_Microphone_IsFarField = DEFINE_PROPERTYKEY(
    0x8943B373,
    0x388C,
    0x4395,
    0xB5,
    0x57,
    0xBC,
    0x6D,
    0xBA,
    0xFF,
    0xAF,
    0xDB,
    6
)

# Name:     System.Devices.AudioDevice.Microphone.SensitivityInDbfs --
# PKEY_Devices_AudioDevice_Microphone_SensitivityInDbfs
# Type:     Double -- VT_R8
# FormatID: {8943B373-388C-4395-B557-BC6DBAFFAFDB}, 3
# Sensitivity information in dBFS for a microphone device.
PKEY_Devices_AudioDevice_Microphone_SensitivityInDbfs = DEFINE_PROPERTYKEY(
    0x8943B373,
    0x388C,
    0x4395,
    0xB5,
    0x57,
    0xBC,
    0x6D,
    0xBA,
    0xFF,
    0xAF,
    0xDB,
    3
)

# Name:     System.Devices.AudioDevice.Microphone.SensitivityInDbfs2 --
# PKEY_Devices_AudioDevice_Microphone_SensitivityInDbfs2
# Type:     Double -- VT_R8
# FormatID: {8943B373-388C-4395-B557-BC6DBAFFAFDB}, 5
# Sensitivity information in dBFS for a microphone device,
# measured after fixed hardware gain (if available). Assumes 0dB software gain
PKEY_Devices_AudioDevice_Microphone_SensitivityInDbfs2 = DEFINE_PROPERTYKEY(
    0x8943B373,
    0x388C,
    0x4395,
    0xB5,
    0x57,
    0xBC,
    0x6D,
    0xBA,
    0xFF,
    0xAF,
    0xDB,
    5
)

# Name:     System.Devices.AudioDevice.Microphone.SignalToNoiseRatioInDb --
# PKEY_Devices_AudioDevice_Microphone_SignalToNoiseRatioInDb
# Type:     Double -- VT_R8
# FormatID: {8943B373-388C-4395-B557-BC6DBAFFAFDB}, 4
# Signal to noise ratio information in dB for a microphone device.
PKEY_Devices_AudioDevice_Microphone_SignalToNoiseRatioInDb = DEFINE_PROPERTYKEY(
    0x8943B373,
    0x388C,
    0x4395,
    0xB5,
    0x57,
    0xBC,
    0x6D,
    0xBA,
    0xFF,
    0xAF,
    0xDB,
    4
)

# Name:     System.Devices.AudioDevice.RawProcessingSupported --
# PKEY_Devices_AudioDevice_RawProcessingSupported
# Type:     Boolean -- VT_BOOL
# FormatID: {8943B373-388C-4395-B557-BC6DBAFFAFDB}, 2
# Raw processing mode support for the audio device.
# If VARIANT_TRUE the device supports raw processing mode.
PKEY_Devices_AudioDevice_RawProcessingSupported = DEFINE_PROPERTYKEY(
    0x8943B373,
    0x388C,
    0x4395,
    0xB5,
    0x57,
    0xBC,
    0x6D,
    0xBA,
    0xFF,
    0xAF,
    0xDB,
    2
)

# Name:     System.Devices.AudioDevice.SpeechProcessingSupported --
# PKEY_Devices_AudioDevice_SpeechProcessingSupported
# Type:     Boolean -- VT_BOOL
# FormatID: {FB1DE864-E06D-47F4-82A6-8A0AEF44493C}, 2
# Speech mode support for the audio device.
# If VARIANT_TRUE the device supports speech mode.
PKEY_Devices_AudioDevice_SpeechProcessingSupported = DEFINE_PROPERTYKEY(
    0xFB1DE864,
    0xE06D,
    0x47F4,
    0x82,
    0xA6,
    0x8A,
    0x0A,
    0xEF,
    0x44,
    0x49,
    0x3C,
    2
)