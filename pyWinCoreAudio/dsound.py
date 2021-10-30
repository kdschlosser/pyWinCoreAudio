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

# Default DirectSound3D algorithm {00000000-0000-0000-0000-000000000000}
DS3DALG_DEFAULT = GUID_NULL

# No virtualization (Pan3D) {C241333F-1C1B-11d2-94F5-00C04FC28ACA}
DS3DALG_NO_VIRTUALIZATION = DEFINE_GUID(0xc241333f, 0x1c1b, 0x11d2, 0x94, 0xf5, 0x0, 0xc0, 0x4f, 0xc2, 0x8a, 0xca)

# High-quality HRTF algorithm {C2413340-1C1B-11d2-94F5-00C04FC28ACA}
DS3DALG_HRTF_FULL = DEFINE_GUID(0xc2413340, 0x1c1b, 0x11d2, 0x94, 0xf5, 0x0, 0xc0, 0x4f, 0xc2, 0x8a, 0xca)

# Lower-quality HRTF algorithm {C2413342-1C1B-11d2-94F5-00C04FC28ACA}
DS3DALG_HRTF_LIGHT = DEFINE_GUID(0xc2413342, 0x1c1b, 0x11d2, 0x94, 0xf5, 0x0, 0xc0, 0x4f, 0xc2, 0x8a, 0xca)