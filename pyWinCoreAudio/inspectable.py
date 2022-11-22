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
import comtypes


IID_IInspectable = MIDL_INTERFACE(
    '{AF86E2E0-B12D-4c6a-9C5A-D7AA65101E90}'
)


class HSTRING__(ctypes.Structure):
    _fields_ = [
        ('unused', INT),
    ]


HSTRING = POINTER(HSTRING__)


class TrustLevel(ENUM):

    BaseTrust = 0
    PartialTrust = BaseTrust + 1
    FullTrust = PartialTrust + 1


class IInspectable(comtypes.IUnknown):
    _case_insensitive_ = False
    _iid_ = IID_IInspectable
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetIids',
            (['out'], POINTER(ULONG), 'iidCount'),
            (['out'], POINTER(POINTER(IID)), 'iids')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetRuntimeClassName',
            (['out'], POINTER(HSTRING), 'className')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetTrustLevel',
            (['out'], POINTER(TrustLevel), 'trustLevel')
        )
    )
