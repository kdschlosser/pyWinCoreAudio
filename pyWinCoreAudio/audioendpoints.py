# -*- coding: utf-8 -*-
#
# This file is part of EventGhost.
# Copyright © 2005-2016 EventGhost Project <http://www.eventghost.net/>
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
import comtypes


IID_IAudioEndpointFormatControl = IID(
    '{784CFD40-9F89-456E-A1A6-873B006A664E}'
)


class IAudioEndpointFormatControl(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IAudioEndpointFormatControl
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'ResetToDefault',
            (['in'], DWORD, 'ResetFlags')
        )
    )


# noinspection PyTypeChecker
PIAudioEndpointFormatControl = POINTER(IAudioEndpointFormatControl)
