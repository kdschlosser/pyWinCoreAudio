import ctypes
import sys
import comtypes.automation
import comtypes

COMMETHOD = comtypes.COMMETHOD
POINTER = ctypes.POINTER
VARIANT = comtypes.automation.VARIANT
VARTYPE = comtypes.automation.VARTYPE
LPVARTYPE = POINTER(VARTYPE)

GUID = comtypes.GUID
LPGUID = POINTER(comtypes.GUID)
LPCGUID = POINTER(GUID)

REFIID = POINTER(GUID)
IID = GUID
CLSID = GUID
REFGUID = POINTER(GUID)


if sys.maxsize > 2**32:
    LONG_PTR = ctypes.c_int64
    LPARAM = ctypes.c_int64
    ULONG_PTR = ctypes.c_uint64
else:
    LONG_PTR = ctypes.c_long
    LPARAM = ctypes.c_long
    ULONG_PTR = ctypes.c_ulong


DWORD = ctypes.c_ulong
DWORDLONG = ctypes.c_uint64
LPDWORD = POINTER(DWORD)

INT = ctypes.c_int
INT8 = ctypes.c_int8
INT16 = ctypes.c_int16
INT32 = ctypes.c_int32
INT64 = ctypes.c_int64

INT_PTR = POINTER(INT)
LPINT = POINTER(INT)


UINT = ctypes.c_uint
UINT8 = ctypes.c_uint8
UINT16 = ctypes.c_uint16
UINT32 = ctypes.c_uint32
UINT64 = ctypes.c_uint64

UINT_PTR = POINTER(UINT)
LPUINT = POINTER(UINT)
LPUINT8 = POINTER(UINT8)
LPUINT16 = POINTER(UINT16)
LPUINT32 = POINTER(UINT32)
LPUINT64 = POINTER(UINT64)


ULONG = ctypes.c_ulong
ULONGLONG = ctypes.c_ulonglong
LPULONG = POINTER(ULONG)

LONG = ctypes.c_long
PLONG = POINTER(LONG)
LONGLONG = ctypes.c_longlong
LPLONG = POINTER(LONG)

FLOAT = ctypes.c_float
FLOAT32 = ctypes.c_float
LPFLOAT = POINTER(FLOAT)

DOUBLE = ctypes.c_double

HANDLE = ctypes.c_void_p
HRESULT = ctypes.c_long
LPHRESULT = POINTER(HRESULT)


VOID = ctypes.c_void_p
PVOID = ctypes.c_void_p
LPVOID = ctypes.c_void_p
LPCVOID = ctypes.c_void_p


BOOL = ctypes.c_int
WINBOOL = BOOL
VARIANT_BOOL = ctypes.c_short
LPBOOL = POINTER(BOOL)

UBYTE = ctypes.c_ubyte
BYTE = ctypes.c_byte
LPBYTE = POINTER(BYTE)

CHAR = ctypes.c_char
UCHAR = ctypes.c_ubyte

SHORT = ctypes.c_short
USHORT = ctypes.c_ushort

WCHAR = ctypes.c_wchar
PWCHAR = POINTER(WCHAR)
LPWCHAR = POINTER(WCHAR)

OLECHAR = ctypes.c_wchar
BSTR = POINTER(OLECHAR)
LPBSTR = POINTER(BSTR)

NULL = None

STRING = ctypes.c_char_p
WSTRING = ctypes.c_wchar_p
LPCWSTR = POINTER(WCHAR)
LPCSTR = ctypes.c_char_p
LPWSTR = POINTER(WCHAR)
LPSTR = POINTER(CHAR)

SCODE = ctypes.c_long

HNSTIME = ctypes.c_longlong
LPHNSTIME = POINTER(HNSTIME)

REFERENCE_TIME = ctypes.c_longlong
LPREFERENCE_TIME = POINTER(REFERENCE_TIME)

WORD = ctypes.c_ushort

COLORREF = DWORD

SIZE_T = ctypes.c_size_t

ACCESS_MASK = DWORD

class _FILETIME(ctypes.Structure):
    _fields_ = [
        ('dwLowDateTime', DWORD),
        ('dwHighDateTime', DWORD)
    ]


FILETIME = _FILETIME


class _LARGE_INTEGER(ctypes.Structure):
    class union(ctypes.Union):
        class struct(ctypes.Structure):
            _fields_ = [
                ('LowPart', ULONG),
                ('HighPart', LONG)
            ]

        _anonymous_ = ('struct',)
        _fields_ = [
            ('struct', struct),
            ('QuadPart', INT64)
        ]

    _anonymous_ = ('union',)
    _fields_ = [('union', union)]


LARGE_INTEGER = _LARGE_INTEGER


class _ULARGE_INTEGER(ctypes.Structure):
    class union(ctypes.Union):
        class struct(ctypes.Structure):
            _fields_ = [
                ('LowPart', ULONG),
                ('HighPart', ULONG)
            ]

        _anonymous_ = ('struct',)
        _fields_ = [
            ('struct', struct),
            ('QuadPart', UINT64)
        ]

    _anonymous_ = ('union',)
    _fields_ = [('union', union)]


ULARGE_INTEGER = _ULARGE_INTEGER


# the following code is something I will usually avoid
# but seeing as how there are so many enumerations I figured
# what the hell.
# This code is alot of smoke and mirrors or as I call it "voodoo magic code"
# The enumarations are used quite often as return values to identify various
# parts and pieces of a sound device and endpoints and I wanted a mchanism to be
# able to attach a description to the integer values contained within
# the enumerations. One of the nice things the developers over at Microsoft
# did was they gave pretty good names top the enumeration values and I thought
# if I could use those value names or portions of them as the descriptions
# that would save me quite a bit of time having to hard code all of those
# descriptions. Well the code  below is the result of that effort and it works!
# well mostly. I did give myself a way to also hard code the descriptions to
# handle the oddities that occur.

# I created a wraper around the `int` class that allows me to pass 2 arguments to the
# constructor. the first argument is the integer and the second is the description string.
# In python 2 this was easily accomplished by creating the wrapper and overriding the
# __init__ method and calling super and only passing the integer to it. This cannot be
# done in python 3 because the value does not get set in the __init__ method. It gets
# set somewhere else, I have not been able to locate where. To fix that issue I created
# a metaclass which sits in the middle. It intercepts the call to the __init__ method
# and then create the class only passing the ineteger and set the description after the
# instance has been created.
class EnumValueMeta(type):

    def __call__(cls, value, description):
        instance = super(EnumValueMeta, cls).__call__(value)
        instance.description = description

        return instance


# this is the int wrapper class, I only overrided the __init__
# for keeping an IDE happy with passing the description.
# I overrided the __str__ method so it can be used as a label or printed
# and it will return the description.
class ENUM_VALUE(int, metaclass=EnumValueMeta):
    description = None

    def __init__(self, *_, **__):
        super(ENUM_VALUE, self).__init__()

    def __str__(self):
        if self.description is None:
            return int.__str__(self)

        return self.description


# this is the voodoo magic part. The metaclass __init__ gets called when the class gets built
# what is passed to the __init__ is the name of the class, bases (parent classes) and a dictonary
# with the class attributes set in it. I am not able to place anything directly into the dictonary
# nor can I create the class and then set anything in the classes dictonary. I can however use setattr
# to set what I need to. I parse the class name and build a prefix if the prefix is found in the
# begining of the attribute name I strip it off. Then I parse whatever is left over of the attribute
# name to add spaces where they need to be added. place those into my fancy ENUM_VALUE class and attach
# that instance to the attribute name using setattr.

# after all is said and done this is the output
# AUDCLNT_SHAREMODE
#     AUDCLNT_SHAREMODE_SHARED = Shared = 1
#     AUDCLNT_SHAREMODE_EXCLUSIVE = Exclusive = 2
# AUDIO_STREAM_CATEGORY
#     AudioCategory_Other = Other = 0
#     AudioCategory_ForegroundOnlyMedia = Foreground Only Media = 1
#     AudioCategory_BackgroundCapableMedia = Background Capable Media = 2
#     AudioCategory_Communications = Communications = 3
#     AudioCategory_Alerts = Alerts = 4
#     AudioCategory_SoundEffects = Sound Effects = 5
#     AudioCategory_GameEffects = Game Effects = 6
#     AudioCategory_GameMedia = Game Media = 7
#     AudioCategory_GameChat = Game Chat = 8
#     AudioCategory_Speech = Speech = 9
#     AudioCategory_Movie = Movie = 10
#     AudioCategory_Media = Media = 11
# STGM
#     STGM_READ = Read = 0
# AUDCLNT_BUFFERFLAGS
#     AUDCLNT_BUFFERFLAGS_DATA_DISCONTINUITY = Data Discontinuity = 1
#     AUDCLNT_BUFFERFLAGS_SILENT = Silent = 2
#     AUDCLNT_BUFFERFLAGS_TIMESTAMP_ERROR = Timestamp Error = 4
# AUDCLNT_STREAMOPTIONS
#     AUDCLNT_STREAMOPTIONS_NONE = None = 0
#     AUDCLNT_STREAMOPTIONS_RAW = Raw = 1
#     AUDCLNT_STREAMOPTIONS_MATCH_FORMAT = Match Format = 2
class EnumMeta(type(INT)):

    def __init__(cls, name, bases, dct):
        super(EnumMeta, cls).__init__(name, bases, dct)

        replace_name = None
        temp_name = ''
        last_char = ''

        for char in name:
            if char.isupper() and not last_char.isupper() and temp_name:
                temp_name += '_'

            temp_name += char
            last_char = char

        for key in list(cls.__dict__.keys())[:]:
            if key.startswith('_'):
                continue

            value = cls.__dict__[key]
            if isinstance(value, ENUM_VALUE):
                continue
            if not isinstance(value, int):
                continue

            if replace_name is None:
                split_name = temp_name.split('_')

                def check_name(n1, n2):
                    return key.upper().startswith((n1 + n2).upper())

                replace_name = ''

                for nm in split_name:
                    if check_name(replace_name, nm):
                        replace_name += nm

                replace_name2 = key[0]
                index = 1
                while name.upper().startswith(replace_name2.upper()):
                    try:
                        replace_name2 += key[index]
                        index += 1
                    except IndexError:
                        replace_name2 = ''
                        break

                if len(replace_name2) > 0:
                    replace_name2 = replace_name2[:-1]

                    if len(replace_name2) > len(replace_name):
                        replace_name = replace_name2

                if replace_name:
                    replace_name = key[:len(replace_name)]
                else:
                    replace_name = None

            new_d = ''
            last_char = ''

            if replace_name is not None and key.startswith(replace_name):
                old_d = key.replace(replace_name, '', 1)

            else:
                old_d = key

            if old_d.startswith('_'):
                old_d = old_d[1:]

            for char in old_d:
                if char == '_':
                    if last_char.isupper():
                        new_d += ' '
                    elif last_char.isdigit():
                        new_d += '_'

                    continue

                if char.isupper() and new_d and not last_char.isupper():
                    new_d += ' '
                new_d += char
                last_char = char

            value = ENUM_VALUE(value, new_d.title())
            setattr(cls, key, value)


class ENUM(INT, metaclass=EnumMeta):

    @classmethod
    def get(cls, val):
        for value in cls.__dict__.values():
            if val in (value, str(value)):
                return value

        return val
