# -*- coding: utf-8 -*-
#
# This file is part of EventGhost.
# Copyright © 2005-2021 EventGhost Project <http://www.eventghost.net/>
#
# EventGhost is free software: you can redistribute it and/or modify it under
# the terms of the GNU General Public License as published by the Free
# Software Foundation, either version 2 of the License, or (at your option)
# any later version.
#
# EventGhost is distributed in the hope that it will be useful, but WITHOUT
# ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
# FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
# more details.
#
# You should have received a copy of the GNU General Public License along
# with EventGhost. If not, see <http://www.gnu.org/licenses/>.

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
