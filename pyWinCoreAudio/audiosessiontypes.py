from .data_types import *


class AudioSessionState(ENUM):
    AudioSessionStateInactive = 0
    AudioSessionStateActive = 1
    AudioSessionStateExpired = 2


PAudioSessionState = POINTER(AudioSessionState)


class AUDCLNT_SHAREMODE(ENUM):
    AUDCLNT_SHAREMODE_SHARED = 1
    AUDCLNT_SHAREMODE_EXCLUSIVE = 2


PAUDCLNT_SHAREMODE = POINTER(AUDCLNT_SHAREMODE)


class AUDIO_STREAM_CATEGORY(ENUM):
    AudioCategory_Other = 0
    AudioCategory_ForegroundOnlyMedia = 1
    AudioCategory_BackgroundCapableMedia = 2
    AudioCategory_Communications = 3
    AudioCategory_Alerts = 4
    AudioCategory_SoundEffects = 5
    AudioCategory_GameEffects = 6
    AudioCategory_GameMedia = 7
    AudioCategory_GameChat = 8
    AudioCategory_Speech = 9
    AudioCategory_Movie = 10
    AudioCategory_Media = 11


PAUDIO_STREAM_CATEGORY = POINTER(AUDIO_STREAM_CATEGORY)
