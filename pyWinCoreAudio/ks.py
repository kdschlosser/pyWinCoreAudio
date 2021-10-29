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
import ctypes


KSPROPERTY_TYPE_GET = 0x00000001
KSPROPERTY_TYPE_SET = 0x00000002
KSPROPERTY_TYPE_SETSUPPORT = 0x00000100
KSPROPERTY_TYPE_BASICSUPPORT = 0x00000200
KSPROPERTY_TYPE_RELATIONS = 0x00000400
KSPROPERTY_TYPE_SERIALIZESET = 0x00000800
KSPROPERTY_TYPE_UNSERIALIZESET = 0x00001000
KSPROPERTY_TYPE_SERIALIZERAW = 0x00002000
KSPROPERTY_TYPE_UNSERIALIZERAW = 0x00004000
KSPROPERTY_TYPE_SERIALIZESIZE = 0x00008000
KSPROPERTY_TYPE_DEFAULTVALUES = 0x00010000
KSPROPERTY_TYPE_TOPOLOGY = 0x10000000
KSPROPERTY_TYPE_HIGHPRIORITY = 0x08000000
KSPROPERTY_TYPE_FSFILTERSCOPE = 0x40000000
KSPROPERTY_TYPE_COPYPAYLOAD = 0x80000000


class KSPROPERTY(ctypes.Structure):
    _fields_ = [
        ('Set', GUID),
        ('Id', ULONG),
        ('Flags', ULONG)
    ]


PKSPROPERTY = POINTER(KSPROPERTY)


class KSMETHOD(ctypes.Structure):
    _fields_ = [
        ('Set', GUID),
        ('Id', ULONG),
        ('Flags', ULONG)
    ]


PKSMETHOD = POINTER(KSMETHOD)


class KSEVENT(ctypes.Structure):
    _fields_ = [
        ('Set', GUID),
        ('Id', ULONG),
        ('Flags', ULONG)
    ]


PKSEVENT = POINTER(KSEVENT)


class KSDATAFORMAT(ctypes.Structure):
    """
    At the minimum, a data format is specified by the MajorFormat, the SubFormat, and the Specifier members.
    A family of similar data formats can share the same values for MajorFormat, SubFormat, and Specifier.
    In that case, the specific data format is distinguished by additional data that follows the
    Specifier member in memory.
    """
    _fields_ = [
        ('FormatSize', ULONG),
        ('Flags', ULONG),
        ('SampleSize', ULONG),
        ('Reserved', ULONG),
        ('MajorFormat', GUID),
        ('SubFormat', GUID),
        ('Specifier', GUID)
    ]

    def __init__(self, *args, **kwargs):

        ctypes.Structure.__init__(self, *args, **kwargs)

        self.FormatSize = ctypes.sizeof(KSDATAFORMAT)

    @property
    def flags(self):
        """
        Set flags to KSDATAFORMAT_ATTRIBUTES (0x2) to indicate that the KSDATAFORMAT is followed in
        memory by a KSMULTIPLE_ITEM of KSATTRIBUTE structures.
        """
        return self.Flags.value

    @flags.setter
    def flags(self, value):
        # noinspection PyAttributeOutsideInit
        self.Flags = value

    @property
    def sample_size(self):
        """
        Specifies the sample size of the data, for fixed sample sizes, or zero,
        if the format has a variable sample size.
        """
        return self.SampleSize.value

    @sample_size.setter
    def sample_size(self, value):
        # noinspection PyAttributeOutsideInit
        self.SampleSize = value

    @property
    def major_format(self):
        """
        Specifies the general format type.

        The data formats that are currently supported can be found in the
        KSDATAFORMAT_TYPE_XXX symbolic constants in the ksmedia.h header file
        that is included in the Windows Driver Kit (WDK).

        A data stream that has no particular format should use
        KSDATAFORMAT_TYPE_STREAM (defined in ks.h) as the value for its
        MajorFormat
        """
        return self.MajorFormat

    @major_format.setter
    def major_format(self, value):
        if not isinstance(value, GUID):
            value = GUID(value)

        # noinspection PyAttributeOutsideInit
        self.MajorFormat = value

    @property
    def sub_format(self):
        """
        Specifies the subformat of a general format type.

        The data subformats that are currently supported can be found in the
        KSDATAFORMAT_SUBTYPE_XXX symbolic constants in the ksmedia.h header
        file that is included in the WDK.

        Major formats that do not support subformats should use the
        KSDATAFORMAT_SUBTYPE_NONE value for this member.
        """
        return self.SubFormat

    @sub_format.setter
    def sub_format(self, value):
        if not isinstance(value, GUID):
            value = GUID(value)

        # noinspection PyAttributeOutsideInit
        self.SubFormat = value

    @property
    def specifier(self):
        """
        Specifies additional data format type information for a specific setting of MajorFormat and SubFormat.

        The significance of this field is determined by the major format (and subformat, if the
        major format supports subformats). For example, Specifier can represent a particular
        encoding of a subformat, or it can be used to specify what type of data structure
        follows KSDATAFORMAT in memory.

        The following specifiers (defined in ks.h) are of general use:

        KSDATAFORMAT_SPECIFIER_NONE: Stands for no specifier. Used for formats that do not support specifiers.
        KSDATAFORMAT_SPECIFIER_FILENAME: Indicates that a null-terminated Unicode string immediately
        follows the KSDATAFORMAT structure in memory.
        KSDATAFORMAT_SPECIFIER_FILEHANDLE: Indicates that a file handle immediately follows KSDATAFORMAT in memory.
        """
        return self.Specifier

    @specifier.setter
    def specifier(self, value):
        if not isinstance(value, GUID):
            value = GUID(value)

        # noinspection PyAttributeOutsideInit
        self.Specifier = value


PKSDATAFORMAT = POINTER(KSDATAFORMAT)
