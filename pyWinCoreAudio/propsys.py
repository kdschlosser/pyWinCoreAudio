
from .data_types import *
from .propidl import (
    PROPVARIANT,
    VARTYPE
)

import ctypes
from comtypes import (
    COMMETHOD,
    helpstring,
    IUnknown,
    CoClass
)

REFPROPVARIANT = POINTER(PROPVARIANT)


class tagCONDITION_OPERATION(ENUM):
    COP_IMPLICIT = 0
    COP_EQUAL = 1
    COP_NOTEQUAL = 2
    COP_LESSTHAN = 3
    COP_GREATERTHAN = 4
    COP_LESSTHANOREQUAL = 5
    COP_GREATERTHANOREQUAL = 6
    COP_VALUE_STARTSWITH = 7
    COP_VALUE_ENDSWITH = 8
    COP_VALUE_CONTAINS = 9
    COP_VALUE_NOTCONTAINS = 10
    COP_DOSWILDCARDS = 11
    COP_WORD_EQUAL = 12
    COP_WORD_STARTSWITH = 13
    COP_APPLICATION_SPECIFIC = 14


CONDITION_OPERATION = tagCONDITION_OPERATION


class tagSHCOLSTATE(ENUM):
    SHCOLSTATE_DEFAULT = 0
    SHCOLSTATE_TYPE_STR = 0x1
    SHCOLSTATE_TYPE_INT = 0x2
    SHCOLSTATE_TYPE_DATE = 0x3
    SHCOLSTATE_TYPEMASK = 0xf
    SHCOLSTATE_ONBYDEFAULT = 0x10
    SHCOLSTATE_SLOW = 0x20
    SHCOLSTATE_EXTENDED = 0x40
    SHCOLSTATE_SECONDARYUI = 0x80
    SHCOLSTATE_HIDDEN = 0x100
    SHCOLSTATE_PREFER_VARCMP = 0x200
    SHCOLSTATE_PREFER_FMTCMP = 0x400
    SHCOLSTATE_NOSORTBYFOLDERNESS = 0x800
    SHCOLSTATE_VIEWONLY = 0x10000
    SHCOLSTATE_BATCHREAD = 0x20000
    SHCOLSTATE_NO_GROUPBY = 0x40000
    SHCOLSTATE_FIXED_WIDTH = 0x1000
    SHCOLSTATE_NODPISCALE = 0x2000
    SHCOLSTATE_FIXED_RATIO = 0x4000
    SHCOLSTATE_DISPLAYMASK = 0xf000


SHCOLSTATE = tagSHCOLSTATE
SHCOLSTATEF = tagSHCOLSTATE


class tagSERIALIZEDPROPSTORAGE(ctypes.Structure):
    pass


SERIALIZEDPROPSTORAGE = tagSERIALIZEDPROPSTORAGE

propsys = ctypes.windll.PROPSYS

LIBID_PropSysObjects = IID('{2cda3294-6c4f-4020-b161-27c530c81fa6}')


class PropSysObjects(object):
    name = 'PropSysObjects'
    _reg_typelib_ = (LIBID_PropSysObjects, 1, 0)


# Forward Declarations

IID_IInitializeWithFile = IID("{B7D14566-0509-4CCE-A71F-0A554233BD9B}")


class IInitializeWithFile(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IInitializeWithFile
    _idlflags_ = []


IID_IInitializeWithStream = IID("{B824B49D-22AC-4161-AC8A-9916E8FA3F7F}")


class IInitializeWithStream(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IInitializeWithStream
    _idlflags_ = []


IID_IPropertyStore = IID("{886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99}")


class IPropertyStore(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertyStore
    _idlflags_ = []

    def SetValue(self, key, value, vt=None):

        prop_var = PROPVARIANT()
        if vt is not None:
            prop_var.vt = vt

        prop_var.value = value

        self.__com_SetValue(key, ctypes.byref(prop_var))

    def GetValue(self, key, return_propvar=False):
        # noinspection PyUnresolvedReferences
        prop_var = PROPVARIANT()

        self.__com_GetValue(key, ctypes.byref(prop_var))

        if return_propvar:
            return prop_var

        return prop_var.value


PIPropertyStore = POINTER(IPropertyStore)


IID_INamedPropertyStore = IID("{71604B0F-97B0-4764-8577-2F13E98A1422}")


class INamedPropertyStore(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_INamedPropertyStore
    _idlflags_ = []


IID_IObjectWithPropertyKey = IID("{FC0CA0A7-C316-4FD2-9031-3E628E6D4F23}")


class IObjectWithPropertyKey(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IObjectWithPropertyKey
    _idlflags_ = []


IID_IPropertyChange = IID("{F917BC8A-1BBA-4478-A245-1BDE03EB9431}")


class IPropertyChange(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertyChange
    _idlflags_ = []


IID_IPropertyChangeArray = IID("{380F5CAD-1B5E-42F2-805D-637FD392D31E}")


class IPropertyChangeArray(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertyChangeArray
    _idlflags_ = []


IID_IPropertyStoreCapabilities = IID("{C8E2D566-186E-4D49-BF41-6909EAD56ACC}")


class IPropertyStoreCapabilities(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertyStoreCapabilities
    _idlflags_ = []


IID_IPropertyStoreCache = IID("{3017056D-9A91-4E90-937D-746C72ABBF4F}")


class IPropertyStoreCache(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertyStoreCache
    _idlflags_ = []


IID_IPropertyEnumType = IID("{11E1FBF9-2D56-4A6B-8DB3-7CD193A471F2}")


class IPropertyEnumType(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertyEnumType
    _idlflags_ = []


IID_IPropertyEnumType2 = IID("{9B6E051C-5DDD-4321-9070-FE2ACB55E794}")


class IPropertyEnumType2(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertyEnumType2
    _idlflags_ = []


IID_IPropertyEnumTypeList = IID("{A99400F4-3D84-4557-94BA-1242FB2CC9A6}")


class IPropertyEnumTypeList(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertyEnumTypeList
    _idlflags_ = []


IID_IPropertyDescription = IID("{6F79D558-3E96-4549-A1D1-7D75D2288814}")


class IPropertyDescription(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertyDescription
    _idlflags_ = []


IID_IPropertyDescription2 = IID("{57D2EDED-5062-400E-B107-5DAE79FE57A6}")


class IPropertyDescription2(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertyDescription2
    _idlflags_ = []


IID_IPropertyDescriptionAliasInfo = IID(
    "{F67104FC-2AF9-46FD-B32D-243C1404F3D1}"
)


class IPropertyDescriptionAliasInfo(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertyDescriptionAliasInfo
    _idlflags_ = []


IID_IPropertyDescriptionSearchInfo = IID(
    "{078F91BD-29A2-440F-924E-46A291524520}"
)


class IPropertyDescriptionSearchInfo(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertyDescriptionSearchInfo
    _idlflags_ = []


IID_IPropertyDescriptionRelatedPropertyInfo = IID(
    "{507393F4-2A3D-4A60-B59E-D9C75716C2DD}"
)


class IPropertyDescriptionRelatedPropertyInfo(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertyDescriptionRelatedPropertyInfo
    _idlflags_ = []


IID_IPropertySystem = IID("{CA724E8A-C3E6-442B-88A4-6FB0DB8035A3}")


class IPropertySystem(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertySystem
    _idlflags_ = []


IID_IPropertyDescriptionList = IID("{1F9FC1D0-C39B-4B26-817F-011967D3440E}")


class IPropertyDescriptionList(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertyDescriptionList
    _idlflags_ = []


IID_IPropertyStoreFactory = IID("{BC110B6D-57E8-4148-A9C6-91015AB2F3A5}")


class IPropertyStoreFactory(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertyStoreFactory
    _idlflags_ = []


IID_IDelayedPropertyStoreFactory = IID(
    "{40D4577F-E237-4BDB-BD69-58F089431B6A}"
)


class IDelayedPropertyStoreFactory(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IDelayedPropertyStoreFactory
    _idlflags_ = []


IID_IPersistSerializedPropStorage = IID(
    "{E318AD57-0AA0-450F-ACA5-6FAB7103D917}"
)


class IPersistSerializedPropStorage(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPersistSerializedPropStorage
    _idlflags_ = []


IID_IPersistSerializedPropStorage2 = IID(
    "{77EFFA68-4F98-4366-BA72-573B3D880571}"
)


class IPersistSerializedPropStorage2(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPersistSerializedPropStorage2
    _idlflags_ = []


IID_IPropertySystemChangeNotify = IID("{FA955FD9-38BE-4879-A6CE-824CF52D609F}")


class IPropertySystemChangeNotify(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertySystemChangeNotify
    _idlflags_ = []


IID_ICreateObject = IID("{75121952-E0D0-43E5-9380-1D80483ACF72}")


class ICreateObject(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_ICreateObject
    _idlflags_ = []


CLSID_InMemoryPropertyStore = IID('{9a02e012-6303-4e1e-b9a1-630f802592c5}')


class InMemoryPropertyStore(CoClass):
    _reg_clsid_ = CLSID_InMemoryPropertyStore
    _idlflags_ = []
    _reg_typelib_ = (LIBID_PropSysObjects, 1, 0)
    _com_interfaces_ = [IPropertyStore]


CLSID_InMemoryPropertyStoreMarshalByValue = IID(
    '{D4CA0E2D-6DA7-4b75-A97C-5F306F0EAEDC}'
)


class InMemoryPropertyStoreMarshalByValue(CoClass):
    _reg_clsid_ = CLSID_InMemoryPropertyStoreMarshalByValue
    _idlflags_ = []
    _reg_typelib_ = (LIBID_PropSysObjects, 1, 0)
    _com_interfaces_ = [IPropertyStore]


CLSID_PropertySystem = IID('{b8967f85-58ae-4f46-9fb2-5d7904798f4b}')


class PropertySystem(CoClass):
    _reg_clsid_ = CLSID_PropertySystem
    _idlflags_ = []
    _reg_typelib_ = (LIBID_PropSysObjects, 1, 0)
    _com_interfaces_ = [IPropertyStore]


REFPROPERTYKEY = POINTER(PROPERTYKEY)


IInitializeWithFile._methods_ = [
    COMMETHOD(
        [helpstring('Method Initialize')],
        HRESULT,
        'Initialize',
        (['in'], LPCWSTR, 'pszFilePath'),
        (['in'], DWORD, 'grfMode'),
    ),
]


IInitializeWithStream._methods_ = [
    COMMETHOD(
        [helpstring('Method Initialize'), 'local'],
        HRESULT,
        'Initialize',
        (['in'], POINTER(IUnknown), 'pstream'),  # IStream interface
        (['in'], DWORD, 'grfMode'),
    ),
]


IPropertyStore._methods_ = [
    COMMETHOD(
        [helpstring('Method GetCount')],
        HRESULT,
        'GetCount',
        (['out'], POINTER(DWORD), 'cProps'),
    ),
    COMMETHOD(
        [helpstring('Method GetAt')],
        HRESULT,
        'GetAt',
        (['in'], DWORD, 'iProp'),
        (['out'], POINTER(PROPERTYKEY), 'pkey'),
    ),
    COMMETHOD(
        [helpstring('Method GetValue')],
        HRESULT,
        'GetValue',
        (['in'], REFPROPERTYKEY, 'key'),
        (['out'], POINTER(PROPVARIANT), 'pv'),
    ),
    COMMETHOD(
        [helpstring('Method SetValue')],
        HRESULT,
        'SetValue',
        (['in'], REFPROPERTYKEY, 'key'),
        (['in'], REFPROPVARIANT, 'propvar'),
    ),
    COMMETHOD(
        [helpstring('Method Commit')],
        HRESULT,
        'Commit',
    ),
]

LPPROPERTYSTORE = POINTER(IPropertyStore)

# HRESULT PropVariantToWinRTPropertyValue(
#     _In_ REFPROPVARIANT propvar, 
#     _In_ REFIID riid, 
#     _COM_Outptr_result_maybenull_ VOID **ppv
# );
PropVariantToWinRTPropertyValue = (
    propsys.PropVariantToWinRTPropertyValue
)
PropVariantToWinRTPropertyValue.restype = HRESULT

# HRESULT WinRTPropertyValueToPropVariant(
#     _In_opt_ IUnknown *punkPropertyValue, 
#     _Out_ PROPVARIANT *ppropvar
# );
WinRTPropertyValueToPropVariant = (
    propsys.WinRTPropertyValueToPropVariant
)
WinRTPropertyValueToPropVariant.restype = HRESULT


INamedPropertyStore._methods_ = [
    COMMETHOD(
        [helpstring('Method GetNamedValue')],
        HRESULT,
        'GetNamedValue',
        (['in'], LPCWSTR, 'pszName'),
        (['out'], POINTER(PROPVARIANT), 'ppropvar'),
    ),
    COMMETHOD(
        [helpstring('Method SetNamedValue')],
        HRESULT,
        'SetNamedValue',
        (['in'], LPCWSTR, 'pszName'),
        (['in'], REFPROPVARIANT, 'propvar'),
    ),
    COMMETHOD(
        [helpstring('Method GetNameCount')],
        HRESULT,
        'GetNameCount',
        (['out'], POINTER(DWORD), 'pdwCount'),
    ),
    COMMETHOD(
        [helpstring('Method GetNameAt')],
        HRESULT,
        'GetNameAt',
        (['in'], DWORD, 'iProp'),
        (['out'], POINTER(BSTR), 'pbstrName'),
    ),
]


class GETPROPERTYSTOREFLAGS(ENUM):
    GPS_DEFAULT = 0
    GPS_HANDLERPROPERTIESONLY = 0x1
    GPS_READWRITE = 0x2
    GPS_TEMPORARY = 0x4
    GPS_FASTPROPERTIESONLY = 0x8
    GPS_OPENSLOWITEM = 0x10
    GPS_DELAYCREATION = 0x20
    GPS_BESTEFFORT = 0x40
    GPS_NO_OPLOCK = 0x80
    GPS_PREFERQUERYPROPERTIES = 0x100
    GPS_EXTRINSICPROPERTIES = 0x200
    GPS_EXTRINSICPROPERTIESONLY = 0x400
    GPS_VOLATILEPROPERTIES = 0x800
    GPS_VOLATILEPROPERTIESONLY = 0x1000
    GPS_MASK_VALID = 0x1FFF


IObjectWithPropertyKey._methods_ = [
    COMMETHOD(
        [helpstring('Method SetPropertyKey')],
        HRESULT,
        'SetPropertyKey',
        (['in'], REFPROPERTYKEY, 'key'),
    ),
    COMMETHOD(
        [helpstring('Method GetPropertyKey')],
        HRESULT,
        'GetPropertyKey',
        (['out'], POINTER(PROPERTYKEY), 'pkey'),
    ),
]


class PKA_FLAGS(ENUM):
    PKA_SET = 0
    PKA_APPEND = PKA_SET + 1
    PKA_DELETE = PKA_APPEND + 1


IPropertyChange._methods_ = [
    COMMETHOD(
        [helpstring('Method ApplyToPropVariant')],
        HRESULT,
        'ApplyToPropVariant',
        (['in'], REFPROPVARIANT, 'propvarIn'),
        (['out'], POINTER(PROPVARIANT), 'ppropvarOut'),
    ),
]


IPropertyChangeArray._methods_ = [
    COMMETHOD(
        [helpstring('Method GetCount')],
        HRESULT,
        'GetCount',
        (['out'], POINTER(UINT), 'pcOperations'),
    ),
    COMMETHOD(
        [helpstring('Method GetAt')],
        HRESULT,
        'GetAt',
        (['in'], UINT, 'iIndex'),
        (['in'], REFIID, 'riid'),
        (['out', 'iid_is'], POINTER(POINTER(VOID)), 'ppv'),
    ),
    COMMETHOD(
        [helpstring('Method InsertAt')],
        HRESULT,
        'InsertAt',
        (['in'], UINT, 'iIndex'),
        (['in'], POINTER(IPropertyChange), 'ppropChange'),
    ),
    COMMETHOD(
        [helpstring('Method Append')],
        HRESULT,
        'Append',
        (['in'], POINTER(IPropertyChange), 'ppropChange'),
    ),
    COMMETHOD(
        [helpstring('Method AppendOrReplace')],
        HRESULT,
        'AppendOrReplace',
        (['in'], POINTER(IPropertyChange), 'ppropChange'),
    ),
    COMMETHOD(
        [helpstring('Method RemoveAt')],
        HRESULT,
        'RemoveAt',
        (['in'], UINT, 'iIndex'),
    ),
    COMMETHOD(
        [helpstring('Method IsKeyInArray')],
        HRESULT,
        'IsKeyInArray',
        (['in'], REFPROPERTYKEY, 'key'),
    ),
]


IPropertyStoreCapabilities._methods_ = [
    COMMETHOD(
        [helpstring('Method IsPropertyWritable')],
        HRESULT,
        'IsPropertyWritable',
        (['in'], REFPROPERTYKEY, 'key'),
    ),
]


class PSC_STATE(ENUM):
    PSC_NORMAL = 0
    PSC_NOTINSOURCE = 1
    PSC_DIRTY = 2
    PSC_READONLY = 3


IPropertyStoreCache._methods_ = [
    COMMETHOD(
        [helpstring('Method GetState')],
        HRESULT,
        'GetState',
        (['in'], REFPROPERTYKEY, 'key'),
        (['out'], POINTER(PSC_STATE), 'pstate'),
    ),
    COMMETHOD(
        [helpstring('Method GetValueAndState')],
        HRESULT,
        'GetValueAndState',
        (['in'], REFPROPERTYKEY, 'key'),
        (['out'], POINTER(PROPVARIANT), 'ppropvar'),
        (['out'], POINTER(PSC_STATE), 'pstate'),
    ),
    COMMETHOD(
        [helpstring('Method SetState')],
        HRESULT,
        'SetState',
        (['in'], REFPROPERTYKEY, 'key'),
        (['in'], PSC_STATE, 'state'),
    ),
    COMMETHOD(
        [helpstring('Method SetValueAndState')],
        HRESULT,
        'SetValueAndState',
        (['in'], REFPROPERTYKEY, 'key'),
        (['unique', 'in'], POINTER(PROPVARIANT), 'ppropvar'),
        (['in'], PSC_STATE, 'state'),
    ),
]


class PROPENUMTYPE(ENUM):
    PET_DISCRETEVALUE = 0
    PET_RANGEDVALUE = 1
    PET_DEFAULTVALUE = 2
    PET_ENDRANGE = 3


IPropertyEnumType._methods_ = [
    COMMETHOD(
        [helpstring('Method GetEnumType')],
        HRESULT,
        'GetEnumType',
        (['out'], POINTER(PROPENUMTYPE), 'penumtype'),
    ),
    COMMETHOD(
        [helpstring('Method GetValue')],
        HRESULT,
        'GetValue',
        (['out'], POINTER(PROPVARIANT), 'ppropvar'),
    ),
    COMMETHOD(
        [helpstring('Method GetRangeMinValue')],
        HRESULT,
        'GetRangeMinValue',
        (['out'], POINTER(PROPVARIANT), 'ppropvarMin'),
    ),
    COMMETHOD(
        [helpstring('Method GetRangeSetValue')],
        HRESULT,
        'GetRangeSetValue',
        (['out'], POINTER(PROPVARIANT), 'ppropvarSet'),
    ),
    COMMETHOD(
        [helpstring('Method GetDisplayText')],
        HRESULT,
        'GetDisplayText',
        (['out'], POINTER(LPWSTR), 'ppszDisplay'),
    ),
]


IPropertyEnumType2._methods_ = [
    COMMETHOD(
        [helpstring('Method GetImageReference')],
        HRESULT,
        'GetImageReference',
        (['out'], POINTER(LPWSTR), 'ppszImageRes'),
    ),
]


IPropertyEnumTypeList._methods_ = [
    COMMETHOD(
        [helpstring('Method GetCount')],
        HRESULT,
        'GetCount',
        (['out'], POINTER(UINT), 'pctypes'),
    ),
    COMMETHOD(
        [helpstring('Method GetAt')],
        HRESULT,
        'GetAt',
        (['in'], UINT, 'itype'),
        (['in'], REFIID, 'riid'),
        (['out', 'iid_is'], POINTER(POINTER(VOID)), 'ppv'),
    ),
    COMMETHOD(
        [helpstring('Method GetConditionAt')],
        HRESULT,
        'GetConditionAt',
        (['in'], UINT, 'nIndex'),
        (['in'], REFIID, 'riid'),
        (['iid_is', 'out'], POINTER(POINTER(VOID)), 'ppv'),
    ),
    COMMETHOD(
        [helpstring('Method FindMatchingIndex')],
        HRESULT,
        'FindMatchingIndex',
        (['in'], REFPROPVARIANT, 'propvarCmp'),
        (['out'], POINTER(UINT), 'pnIndex'),
    ),
]


class PROPDESC_TYPE_FLAGS(ENUM):
    PDTF_DEFAULT = 0
    PDTF_MULTIPLEVALUES = 0x1
    PDTF_ISINNATE = 0x2
    PDTF_ISGROUP = 0x4
    PDTF_CANGROUPBY = 0x8
    PDTF_CANSTACKBY = 0x10
    PDTF_ISTREEPROPERTY = 0x20
    PDTF_INCLUDEINFULLTEXTQUERY = 0x40
    PDTF_ISVIEWABLE = 0x80
    PDTF_ISQUERYABLE = 0x100
    PDTF_CANBEPURGED = 0x200
    PDTF_SEARCHRAWVALUE = 0x400
    PDTF_DONTCOERCEEMPTYSTRINGS = 0x800
    PDTF_ALWAYSINSUPPLEMENTALSTORE = 0x1000
    PDTF_ISSYSTEMPROPERTY = 0x80000000
    PDTF_MASK_ALL = 0x80001FFF


class PROPDESC_VIEW_FLAGS(ENUM):
    PDVF_DEFAULT = 0
    PDVF_CENTERALIGN = 0x1
    PDVF_RIGHTALIGN = 0x2
    PDVF_BEGINNEWGROUP = 0x4
    PDVF_FILLAREA = 0x8
    PDVF_SORTDESCENDING = 0x10
    PDVF_SHOWONLYIFPRESENT = 0x20
    PDVF_SHOWBYDEFAULT = 0x40
    PDVF_SHOWINPRIMARYLIST = 0x80
    PDVF_SHOWINSECONDARYLIST = 0x100
    PDVF_HIDELABEL = 0x200
    PDVF_HIDDEN = 0x800
    PDVF_CANWRAP = 0x1000
    PDVF_MASK_ALL = 0x1BFF


class PROPDESC_DISPLAYTYPE(ENUM):
    PDDT_STRING = 0
    PDDT_NUMBER = 1
    PDDT_BOOLEAN = 2
    PDDT_DATETIME = 3
    PDDT_ENUMERATED = 4


class PROPDESC_GROUPING_RANGE(ENUM):
    PDGR_DISCRETE = 0
    PDGR_ALPHANUMERIC = 1
    PDGR_SIZE = 2
    PDGR_DYNAMIC = 3
    PDGR_DATE = 4
    PDGR_PERCENT = 5
    PDGR_ENUMERATED = 6


class PROPDESC_FORMAT_FLAGS(ENUM):
    PDFF_DEFAULT = 0
    PDFF_PREFIXNAME = 0x1
    PDFF_FILENAME = 0x2
    PDFF_ALWAYSKB = 0x4
    PDFF_RESERVED_RIGHTTOLEFT = 0x8
    PDFF_SHORTTIME = 0x10
    PDFF_LONGTIME = 0x20
    PDFF_HIDETIME = 0x40
    PDFF_SHORTDATE = 0x80
    PDFF_LONGDATE = 0x100
    PDFF_HIDEDATE = 0x200
    PDFF_RELATIVEDATE = 0x400
    PDFF_USEEDITINVITATION = 0x800
    PDFF_READONLY = 0x1000
    PDFF_NOAUTOREADINGORDER = 0x2000


class PROPDESC_SORTDESCRIPTION(ENUM):
    PDSD_GENERAL = 0
    PDSD_A_Z = 1
    PDSD_LOWEST_HIGHEST = 2
    PDSD_SMALLEST_BIGGEST = 3
    PDSD_OLDEST_NEWEST = 4


class PROPDESC_RELATIVEDESCRIPTION_TYPE(ENUM):
    PDRDT_GENERAL = 0
    PDRDT_DATE = 1
    PDRDT_SIZE = 2
    PDRDT_COUNT = 3
    PDRDT_REVISION = 4
    PDRDT_LENGTH = 5
    PDRDT_DURATION = 6
    PDRDT_SPEED = 7
    PDRDT_RATE = 8
    PDRDT_RATING = 9
    PDRDT_PRIORITY = 10


class PROPDESC_AGGREGATION_TYPE(ENUM):
    PDAT_DEFAULT = 0
    PDAT_FIRST = 1
    PDAT_SUM = 2
    PDAT_AVERAGE = 3
    PDAT_DATERANGE = 4
    PDAT_UNION = 5
    PDAT_MAX = 6
    PDAT_MIN = 7


class PROPDESC_CONDITION_TYPE(ENUM):
    PDCOT_NONE = 0
    PDCOT_STRING = 1
    PDCOT_SIZE = 2
    PDCOT_DATETIME = 3
    PDCOT_BOOLEAN = 4
    PDCOT_NUMBER = 5


IPropertyDescription._methods_ = [
    COMMETHOD(
        [helpstring('Method GetPropertyKey')],
        HRESULT,
        'GetPropertyKey',
        (['out'], POINTER(PROPERTYKEY), 'pkey'),
    ),
    COMMETHOD(
        [helpstring('Method GetCanonicalName')],
        HRESULT,
        'GetCanonicalName',
        (['out', 'in'], POINTER(LPWSTR), 'ppszName'),
    ),
    COMMETHOD(
        [helpstring('Method GetPropertyType')],
        HRESULT,
        'GetPropertyType',
        (['out'], POINTER(VARTYPE), 'pvartype'),
    ),
    COMMETHOD(
        [helpstring('Method GetDisplayName')],
        HRESULT,
        'GetDisplayName',
        (['out', 'in'], POINTER(LPWSTR), 'ppszName'),
    ),
    COMMETHOD(
        [helpstring('Method GetEditInvitation')],
        HRESULT,
        'GetEditInvitation',
        (['out', 'in'], POINTER(LPWSTR), 'ppszInvite'),
    ),
    COMMETHOD(
        [helpstring('Method GetTypeFlags')],
        HRESULT,
        'GetTypeFlags',
        (['in'], PROPDESC_TYPE_FLAGS, 'mask'),
        (['out'], POINTER(PROPDESC_TYPE_FLAGS), 'ppdtFlags'),
    ),
    COMMETHOD(
        [helpstring('Method GetViewFlags')],
        HRESULT,
        'GetViewFlags',
        (['out'], POINTER(PROPDESC_VIEW_FLAGS), 'ppdvFlags'),
    ),
    COMMETHOD(
        [helpstring('Method GetDefaultColumnWidth')],
        HRESULT,
        'GetDefaultColumnWidth',
        (['out'], POINTER(UINT), 'pcxChars'),
    ),
    COMMETHOD(
        [helpstring('Method GetDisplayType')],
        HRESULT,
        'GetDisplayType',
        (
            ['out'],
            POINTER(PROPDESC_DISPLAYTYPE),
            'pdisplaytype'
        ),
    ),
    COMMETHOD(
        [helpstring('Method GetColumnState')],
        HRESULT,
        'GetColumnState',
        (['out'], POINTER(SHCOLSTATEF), 'pcsFlags'),
    ),
    COMMETHOD(
        [helpstring('Method GetGroupingRange')],
        HRESULT,
        'GetGroupingRange',
        (['out'], POINTER(PROPDESC_GROUPING_RANGE), 'pgr'),
    ),
    COMMETHOD(
        [helpstring('Method GetRelativeDescriptionType')],
        HRESULT,
        'GetRelativeDescriptionType',
        (
            ['out'],
            POINTER(PROPDESC_RELATIVEDESCRIPTION_TYPE),
            'prdt'
        ),
    ),
    COMMETHOD(
        [helpstring('Method GetRelativeDescription')],
        HRESULT,
        'GetRelativeDescription',
        (['in'], REFPROPVARIANT, 'propvar1'),
        (['in'], REFPROPVARIANT, 'propvar2'),
        (['out', 'in'], POINTER(LPWSTR), 'ppszDesc1'),
        (['out', 'in'], POINTER(LPWSTR), 'ppszDesc2'),
    ),
    COMMETHOD(
        [helpstring('Method GetSortDescription')],
        HRESULT,
        'GetSortDescription',
        (['out'], POINTER(PROPDESC_SORTDESCRIPTION), 'psd'),
    ),
    COMMETHOD(
        [helpstring('Method GetSortDescriptionLabel')],
        HRESULT,
        'GetSortDescriptionLabel',
        (['in'], BOOL, 'fDescending'),
        (['out', 'in'], POINTER(LPWSTR), 'ppszDescription'),
    ),
    COMMETHOD(
        [helpstring('Method GetAggregationType')],
        HRESULT,
        'GetAggregationType',
        (
            ['out'],
            POINTER(PROPDESC_AGGREGATION_TYPE),
            'paggtype'
        ),
    ),
    COMMETHOD(
        [helpstring('Method GetConditionType')],
        HRESULT,
        'GetConditionType',
        (
            ['out'],
            POINTER(PROPDESC_CONDITION_TYPE),
            'pcontype'
        ),
        (['out'], POINTER(CONDITION_OPERATION), 'popDefault'),
    ),
    COMMETHOD(
        [helpstring('Method GetEnumTypeList')],
        HRESULT,
        'GetEnumTypeList',
        (['in'], REFIID, 'riid'),
        (['out', 'iid_is'], POINTER(POINTER(VOID)), 'ppv'),
    ),
    COMMETHOD(
        [helpstring('Method CoerceToCanonicalValue'), 'local'],
        HRESULT,
        'CoerceToCanonicalValue',
        (['out', 'in'], POINTER(PROPVARIANT), 'ppropvar'),
    ),
    COMMETHOD(
        [helpstring('Method FormatForDisplay')],
        HRESULT,
        'FormatForDisplay',
        (['in'], REFPROPVARIANT, 'propvar'),
        (['in'], PROPDESC_FORMAT_FLAGS, 'pdfFlags'),
        (['out', 'in'], POINTER(LPWSTR), 'ppszDisplay'),
    ),
    COMMETHOD(
        [helpstring('Method IsValueCanonical')],
        HRESULT,
        'IsValueCanonical',
        (['in'], REFPROPVARIANT, 'propvar'),
    ),
]


IPropertyDescription2._methods_ = [
    COMMETHOD(
        [helpstring('Method GetImageReferenceForValue')],
        HRESULT,
        'GetImageReferenceForValue',
        (['in'], REFPROPVARIANT, 'propvar'),
        (['out', 'in'], POINTER(LPWSTR), 'ppszImageRes'),
    ),
]


IPropertyDescriptionAliasInfo._methods_ = [
    COMMETHOD(
        [helpstring('Method GetSortByAlias')],
        HRESULT,
        'GetSortByAlias',
        (['in'], REFIID, 'riid'),
        (['iid_is', 'out'], POINTER(POINTER(VOID)), 'ppv'),
    ),
    COMMETHOD(
        [helpstring('Method GetAdditionalSortByAliases')],
        HRESULT,
        'GetAdditionalSortByAliases',
        (['in'], REFIID, 'riid'),
        (['iid_is', 'out'], POINTER(POINTER(VOID)), 'ppv'),
    ),
]


class PROPDESC_SEARCHINFO_FLAGS(ENUM):
    PDSIF_DEFAULT = 0
    PDSIF_ININVERTEDINDEX = 0x1
    PDSIF_ISCOLUMN = 0x2
    PDSIF_ISCOLUMNSPARSE = 0x4
    PDSIF_ALWAYSINCLUDE = 0x8
    PDSIF_USEFORTYPEAHEAD = 0x10


class PROPDESC_COLUMNINDEX_TYPE(ENUM):
    PDCIT_NONE = 0
    PDCIT_ONDISK = 1
    PDCIT_INMEMORY = 2
    PDCIT_ONDEMAND = 3
    PDCIT_ONDISKALL = 4
    PDCIT_ONDISKVECTOR = 5


IPropertyDescriptionSearchInfo._methods_ = [
    COMMETHOD(
        [helpstring('Method GetSearchInfoFlags')],
        HRESULT,
        'GetSearchInfoFlags',
        (
            ['out'],
            POINTER(PROPDESC_SEARCHINFO_FLAGS),
            'ppdsiFlags'
        ),
    ),
    COMMETHOD(
        [helpstring('Method GetColumnIndexType')],
        HRESULT,
        'GetColumnIndexType',
        (
            ['out'],
            POINTER(PROPDESC_COLUMNINDEX_TYPE),
            'ppdciType'
        ),
    ),
    COMMETHOD(
        [helpstring('Method GetProjectionString')],
        HRESULT,
        'GetProjectionString',
        (['out'], POINTER(LPWSTR), 'ppszProjection'),
    ),
    COMMETHOD(
        [helpstring('Method GetMaxSize')],
        HRESULT,
        'GetMaxSize',
        (['out'], POINTER(UINT), 'pcbMaxSize'),
    ),
]


IPropertyDescriptionRelatedPropertyInfo._methods_ = [
    COMMETHOD(
        [helpstring('Method GetRelatedProperty')],
        HRESULT,
        'GetRelatedProperty',
        (['in'], LPCWSTR, 'pszRelationshipName'),
        (['in'], REFIID, 'riid'),
        (['iid_is', 'out'], POINTER(POINTER(VOID)), 'ppv'),
    ),
]


class PROPDESC_ENUMFILTER(ENUM):
    PDEF_ALL = 0
    PDEF_SYSTEM = 1
    PDEF_NONSYSTEM = 2
    PDEF_VIEWABLE = 3
    PDEF_QUERYABLE = 4
    PDEF_INFULLTEXTQUERY = 5
    PDEF_COLUMN = 6


IPropertySystem._methods_ = [
    COMMETHOD(
        [helpstring('Method GetPropertyDescription')],
        HRESULT,
        'GetPropertyDescription',
        (['in'], REFPROPERTYKEY, 'propkey'),
        (['in'], REFIID, 'riid'),
        (['iid_is', 'out'], POINTER(POINTER(VOID)), 'ppv'),
    ),
    COMMETHOD(
        [helpstring('Method GetPropertyDescriptionByName')],
        HRESULT,
        'GetPropertyDescriptionByName',
        (['in'], LPCWSTR, 'pszCanonicalName'),
        (['in'], REFIID, 'riid'),
        (['iid_is', 'out'], POINTER(POINTER(VOID)), 'ppv'),
    ),
    COMMETHOD(
        [helpstring('Method GetPropertyDescriptionListFromString')],
        HRESULT,
        'GetPropertyDescriptionListFromString',
        (['in'], LPCWSTR, 'pszPropList'),
        (['in'], REFIID, 'riid'),
        (['out', 'iid_is'], POINTER(POINTER(VOID)), 'ppv'),
    ),
    COMMETHOD(
        [helpstring('Method EnumeratePropertyDescriptions')],
        HRESULT,
        'EnumeratePropertyDescriptions',
        (['in'], PROPDESC_ENUMFILTER, 'filterOn'),
        (['in'], REFIID, 'riid'),
        (['out', 'iid_is'], POINTER(POINTER(VOID)), 'ppv'),
    ),
    COMMETHOD(
        [helpstring('Method FormatForDisplay')],
        HRESULT,
        'FormatForDisplay',
        (['in'], REFPROPERTYKEY, 'key'),
        (['in'], REFPROPVARIANT, 'propvar'),
        (['in'], PROPDESC_FORMAT_FLAGS, 'pdff'),
        (['out', 'in'], LPWSTR, 'pszText'),
        (['range', 'in'], DWORD, 'cchText'),
    ),
    COMMETHOD(
        [helpstring('Method FormatForDisplayAlloc')],
        HRESULT,
        'FormatForDisplayAlloc',
        (['in'], REFPROPERTYKEY, 'key'),
        (['in'], REFPROPVARIANT, 'propvar'),
        (['in'], PROPDESC_FORMAT_FLAGS, 'pdff'),
        (['out', 'in'], POINTER(LPWSTR), 'ppszDisplay'),
    ),
    COMMETHOD(
        [helpstring('Method RegisterPropertySchema')],
        HRESULT,
        'RegisterPropertySchema',
        (['in'], LPCWSTR, 'pszPath'),
    ),
    COMMETHOD(
        [helpstring('Method UnregisterPropertySchema')],
        HRESULT,
        'UnregisterPropertySchema',
        (['in'], LPCWSTR, 'pszPath'),
    ),
    COMMETHOD(
        [helpstring('Method RefreshPropertySchema')],
        HRESULT,
        'RefreshPropertySchema',
    ),
]


IPropertyDescriptionList._methods_ = [
    COMMETHOD(
        [helpstring('Method GetCount')],
        HRESULT,
        'GetCount',
        (['out'], POINTER(UINT), 'pcElem'),
    ),
    COMMETHOD(
        [helpstring('Method GetAt')],
        HRESULT,
        'GetAt',
        (['in'], UINT, 'iElem'),
        (['in'], REFIID, 'riid'),
        (['iid_is', 'out'], POINTER(POINTER(VOID)), 'ppv'),
    ),
]


IPropertyStoreFactory._methods_ = [
    COMMETHOD(
        [helpstring('Method GetPropertyStore')],
        HRESULT,
        'GetPropertyStore',
        (['in'], GETPROPERTYSTOREFLAGS, 'flags'),
        (
            ['unique', 'in'],
            POINTER(IUnknown),
            'pUnkFactory'
        ),
        (['in'], REFIID, 'riid'),
        (['iid_is', 'out'], POINTER(POINTER(VOID)), 'ppv'),
    ),
    COMMETHOD(
        [helpstring('Method GetPropertyStoreForKeys')],
        HRESULT,
        'GetPropertyStoreForKeys',
        (['unique', 'in'], POINTER(PROPERTYKEY), 'rgKeys'),
        (['in'], UINT, 'cKeys'),
        (['in'], GETPROPERTYSTOREFLAGS, 'flags'),
        (['in'], REFIID, 'riid'),
        (['iid_is', 'out'], POINTER(POINTER(VOID)), 'ppv'),
    ),
]


IDelayedPropertyStoreFactory._methods_ = [
    COMMETHOD(
        [helpstring('Method GetDelayedPropertyStore')],
        HRESULT,
        'GetDelayedPropertyStore',
        (['in'], GETPROPERTYSTOREFLAGS, 'flags'),
        (['in'], DWORD, 'dwStoreId'),
        (['in'], REFIID, 'riid'),
        (['iid_is', 'out'], POINTER(POINTER(VOID)), 'ppv'),
    ),
]


class _PERSIST_SPROPSTORE_FLAGS(ENUM):
    FPSPS_DEFAULT = 0
    FPSPS_READONLY = 0x1
    FPSPS_TREAT_NEW_VALUES_AS_DIRTY = 0x2


FPSPS_DEFAULT = _PERSIST_SPROPSTORE_FLAGS.FPSPS_DEFAULT
FPSPS_READONLY = _PERSIST_SPROPSTORE_FLAGS.FPSPS_READONLY
FPSPS_TREAT_NEW_VALUES_AS_DIRTY = (
    _PERSIST_SPROPSTORE_FLAGS.FPSPS_TREAT_NEW_VALUES_AS_DIRTY
)
PERSIST_SPROPSTORE_FLAGS = INT

PUSERIALIZEDPROPSTORAGE = POINTER(SERIALIZEDPROPSTORAGE)
PCUSERIALIZEDPROPSTORAGE = POINTER(SERIALIZEDPROPSTORAGE)


IPersistSerializedPropStorage._methods_ = [
    COMMETHOD(
        [helpstring('Method SetFlags')],
        HRESULT,
        'SetFlags',
        (['in'], PERSIST_SPROPSTORE_FLAGS, 'flags'),
    ),
    COMMETHOD(
        [helpstring('Method SetPropertyStorage')],
        HRESULT,
        'SetPropertyStorage',
        (['in'], PCUSERIALIZEDPROPSTORAGE, 'psps'),
        (['in'], DWORD, 'cb'),
    ),
    COMMETHOD(
        [helpstring('Method GetPropertyStorage')],
        HRESULT,
        'GetPropertyStorage',
        (
            ['out'],
            POINTER(POINTER(SERIALIZEDPROPSTORAGE)),
            'ppsps'
        ),
        (['out'], POINTER(DWORD), 'pcb'),
    ),
]


IPersistSerializedPropStorage2._methods_ = [
    COMMETHOD(
        [helpstring('Method GetPropertyStorageSize')],
        HRESULT,
        'GetPropertyStorageSize',
        (['out'], POINTER(DWORD), 'pcb'),
    ),
    COMMETHOD(
        [helpstring('Method GetPropertyStorageBuffer')],
        HRESULT,
        'GetPropertyStorageBuffer',
        (['out'], POINTER(SERIALIZEDPROPSTORAGE), 'psps'),
        (['in'], DWORD, 'cb'),
        (['out'], POINTER(DWORD), 'pcbWritten'),
    ),
]


IPropertySystemChangeNotify._methods_ = [
    COMMETHOD(
        [helpstring('Method SchemaRefreshed')],
        HRESULT,
        'SchemaRefreshed',
    ),
]

ICreateObject._methods_ = [
    COMMETHOD(
        [helpstring('Method CreateObject')],
        HRESULT,
        'CreateObject',
        (['in'], REFCLSID, 'clsid'),
        (
            ['unique', 'in'],
            POINTER(IUnknown),
            'pUnkOuter'
        ),
        (['in'], REFIID, 'riid'),
        (['out', 'iid_is'], POINTER(POINTER(VOID)), 'ppv'),
    ),
]

# HRESULT PSFormatForDisplay(
# _In_ REFPROPERTYKEY propkey,
# _In_ REFPROPVARIANT propvar,
# _In_ PROPDESC_FORMAT_FLAGS pdfFlags,
# _Out_writes_(cchText) LPWSTR pwszText,
# _In_ DWORD cchText);
PSFormatForDisplay = propsys.PSFormatForDisplay
PSFormatForDisplay.restype = HRESULT

# HRESULT PSFormatForDisplayAlloc(
# _In_ REFPROPERTYKEY key,
# _In_ REFPROPVARIANT propvar,
# _In_ PROPDESC_FORMAT_FLAGS pdff,
# _Outptr_ PWSTR *ppszDisplay);
PSFormatForDisplayAlloc = propsys.PSFormatForDisplayAlloc
PSFormatForDisplayAlloc.restype = HRESULT

# HRESULT PSFormatPropertyValue(
# _In_ IPropertyStore *pps,
# _In_ IPropertyDescription *ppd,
# _In_ PROPDESC_FORMAT_FLAGS pdff,
# _Outptr_ LPWSTR *ppszDisplay);
PSFormatPropertyValue = propsys.PSFormatPropertyValue
PSFormatPropertyValue.restype = HRESULT

# Retrieve the image reference associated with a property value
# (if specified)
# HRESULT PSGetImageReferenceForValue(
# _In_ REFPROPERTYKEY propkey,
# _In_ REFPROPVARIANT propvar,
# _Outptr_ PWSTR *ppszImageRes);
PSGetImageReferenceForValue = propsys.PSGetImageReferenceForValue
PSGetImageReferenceForValue.restype = HRESULT

# Convert a PROPERTYKEY to and from a PWSTR
# HRESULT PSStringFromPropertyKey(
# _In_ REFPROPERTYKEY pkey,
# _Out_writes_(cch) LPWSTR psz,
# _In_ UINT cch);
PSStringFromPropertyKey = propsys.PSStringFromPropertyKey
PSStringFromPropertyKey.restype = HRESULT

# HRESULT PSPropertyKeyFromString(
# _In_ LPCWSTR pszString,
# _Out_ PROPERTYKEY *pkey);
PSPropertyKeyFromString = propsys.PSPropertyKeyFromString
PSPropertyKeyFromString.restype = HRESULT

# Creates an in-memory property store
# Returns an IPropertyStore, IPersistSerializedPropStorage, and
# related interfaces interface

# HRESULT PSCreateMemoryPropertyStore(
# _In_ REFIID riid,
# _Outptr_ VOID **ppv);
PSCreateMemoryPropertyStore = propsys.PSCreateMemoryPropertyStore
PSCreateMemoryPropertyStore.restype = HRESULT

# Create a read-only, delay-bind multiplexing property store
# Returns an IPropertyStore interface or related interfaces

# HRESULT PSCreateDelayedMultiplexPropertyStore(
# _In_ GETPROPERTYSTOREFLAGS flags,
# _In_ IDelayedPropertyStoreFactory *pdpsf,
# _In_reads_(cStores) DWORD *rgStoreIds,
# _In_ DWORD cStores,
# _In_ REFIID riid,
# _Outptr_ VOID **ppv);
PSCreateDelayedMultiplexPropertyStore = (
    propsys.PSCreateDelayedMultiplexPropertyStore
)
PSCreateDelayedMultiplexPropertyStore.restype = HRESULT

# Create a read-only property store from one or more sources
# (which each must support either IPropertyStore or IPropertySetStorage)
#
# Returns an IPropertyStore interface or related interfaces
# HRESULT PSCreateMultiplexPropertyStore(
# _In_reads_(cStores) IUnknown **prgpunkStores,
# _In_ DWORD cStores,
# _In_ REFIID riid,
# _Outptr_ VOID **ppv);
PSCreateMultiplexPropertyStore = propsys.PSCreateMultiplexPropertyStore
PSCreateMultiplexPropertyStore.restype = HRESULT

# Create a container for IPropertyChanges
# Returns an IPropertyChangeArray interface
# HRESULT PSCreatePropertyChangeArray(
# _In_reads_opt_(cChanges) PROPERTYKEY *rgpropkey,
# _In_reads_opt_(cChanges) PKA_FLAGS *rgflags,
# _In_reads_opt_(cChanges) PROPVARIANT *rgpropvar,
# _In_ UINT cChanges,
# _In_ REFIID riid,
# _Outptr_ VOID **ppv);
PSCreatePropertyChangeArray = propsys.PSCreatePropertyChangeArray
PSCreatePropertyChangeArray.restype = HRESULT

# Create a simple property change
# Returns an IPropertyChange interface
# HRESULT PSCreateSimplePropertyChange(
# _In_ PKA_FLAGS flags,
# _In_ REFPROPERTYKEY key,
# _In_ REFPROPVARIANT propvar,
# _In_ REFIID riid,
# _Outptr_ VOID **ppv);
PSCreateSimplePropertyChange = propsys.PSCreateSimplePropertyChange
PSCreateSimplePropertyChange.restype = HRESULT

# Get a property description
# Returns an IPropertyDescription interface
# HRESULT PSGetPropertyDescription(
# _In_ REFPROPERTYKEY propkey,
# _In_ REFIID riid,
# _Outptr_ VOID **ppv);
PSGetPropertyDescription = propsys.PSGetPropertyDescription
PSGetPropertyDescription.restype = HRESULT

# HRESULT PSGetPropertyDescriptionByName(
# _In_ LPCWSTR pszCanonicalName,
# _In_ REFIID riid,
# _Outptr_ VOID **ppv);
PSGetPropertyDescriptionByName = propsys.PSGetPropertyDescriptionByName
PSGetPropertyDescriptionByName.restype = HRESULT

# Lookup a per-machine registered file property handler
# HRESULT PSLookupPropertyHandlerCLSID(
# _In_ PCWSTR pszFilePath,
# _Out_ CLSID *pclsid);
PSLookupPropertyHandlerCLSID = propsys.PSLookupPropertyHandlerCLSID
PSLookupPropertyHandlerCLSID.restype = HRESULT

# Get a property handler, on Vista or downlevel to XP
# punkItem is a shell item created with an SHCreateItemXXX API
# Returns an IPropertyStore
# HRESULT PSGetItemPropertyHandler(
# _In_ IUnknown *punkItem,
# _In_ BOOL fReadWrite,
# _In_ REFIID riid,
# _Outptr_ VOID **ppv);
PSGetItemPropertyHandler = propsys.PSGetItemPropertyHandler
PSGetItemPropertyHandler.restype = HRESULT

# Get a property handler, on Vista or downlevel to XP
# punkItem is a shell item created with an SHCreateItemXXX API
# punkCreateObject supports ICreateObject
# Returns an IPropertyStore
# HRESULT PSGetItemPropertyHandlerWithCreateObject(
# _In_ IUnknown *punkItem,
# _In_ BOOL fReadWrite,
# _In_ IUnknown *punkCreateObject,
# _In_ REFIID riid,
# _Outptr_ VOID **ppv);
PSGetItemPropertyHandlerWithCreateObject = (
    propsys.PSGetItemPropertyHandlerWithCreateObject
)
PSGetItemPropertyHandlerWithCreateObject.restype = HRESULT

# Get or set a property value from a store
# HRESULT PSGetPropertyValue(
# _In_ IPropertyStore *pps,
# _In_ IPropertyDescription *ppd,
# _Out_ PROPVARIANT *ppropvar);
PSGetPropertyValue = propsys.PSGetPropertyValue
PSGetPropertyValue.restype = HRESULT

# HRESULT PSSetPropertyValue(
# _In_ IPropertyStore *pps,
# _In_ IPropertyDescription *ppd,
# _In_ REFPROPVARIANT propvar);
PSSetPropertyValue = propsys.PSSetPropertyValue
PSSetPropertyValue.restype = HRESULT

# Interact with the set of property descriptions
# HRESULT PSRegisterPropertySchema(
# _In_ PCWSTR pszPath);
PSRegisterPropertySchema = propsys.PSRegisterPropertySchema
PSRegisterPropertySchema.restype = HRESULT

# HRESULT PSUnregisterPropertySchema(
# _In_ PCWSTR pszPath);
PSUnregisterPropertySchema = propsys.PSUnregisterPropertySchema
PSUnregisterPropertySchema.restype = HRESULT

# Returns either: IPropertyDescriptionList or IEnumUnknown interfaces
# HRESULT PSEnumeratePropertyDescriptions(
# _In_ PROPDESC_ENUMFILTER filterOn,
# _In_ REFIID riid,
# _Outptr_ VOID **ppv);
PSEnumeratePropertyDescriptions = (
    propsys.PSEnumeratePropertyDescriptions
)
PSEnumeratePropertyDescriptions.restype = HRESULT

# Convert between a PROPERTYKEY and its canonical name
# HRESULT PSGetPropertyKeyFromName(
# _In_ PCWSTR pszName,
# _Out_ PROPERTYKEY *ppropkey);
PSGetPropertyKeyFromName = propsys.PSGetPropertyKeyFromName
PSGetPropertyKeyFromName.restype = HRESULT

# HRESULT PSGetNameFromPropertyKey(
# _In_ REFPROPERTYKEY propkey,
# _Outptr_ PWSTR *ppszCanonicalName);
PSGetNameFromPropertyKey = propsys.PSGetNameFromPropertyKey
PSGetNameFromPropertyKey.restype = HRESULT

# Coerce and canonicalize a property value
# HRESULT PSCoerceToCanonicalValue(
# _In_ REFPROPERTYKEY key,
# _Inout_ PROPVARIANT *ppropvar);
PSCoerceToCanonicalValue = propsys.PSCoerceToCanonicalValue
PSCoerceToCanonicalValue.restype = HRESULT

# Convert a 'prop:' string into a list of property descriptions
# Returns an IPropertyDescriptionList interface
# HRESULT PSGetPropertyDescriptionListFromString(
# _In_ LPCWSTR pszPropList,
# _In_ REFIID riid,
# _Outptr_ VOID **ppv);
PSGetPropertyDescriptionListFromString = (
    propsys.PSGetPropertyDescriptionListFromString
)
PSGetPropertyDescriptionListFromString.restype = HRESULT

# Wrap an IPropertySetStorage interface in an IPropertyStore interface
# Returns an IPropertyStore or related interface
# HRESULT PSCreatePropertyStoreFromPropertySetStorage(
# _In_ IPropertySetStorage *ppss,
# _In_ DWORD grfMode,
# _In_ REFIID riid,
# _Outptr_ VOID **ppv);
PSCreatePropertyStoreFromPropertySetStorage = (
    propsys.PSCreatePropertyStoreFromPropertySetStorage
)
PSCreatePropertyStoreFromPropertySetStorage.restype = HRESULT

# punkSource must support IPropertyStore or IPropertySetStorage
# On success, the returned ppv is guaranteed to support IPropertyStore.
# If punkSource already supports IPropertyStore, no wrapper is created.

# HRESULT PSCreatePropertyStoreFromObject(
# _In_ IUnknown *punk,
# _In_ DWORD grfMode,
# _In_ REFIID riid,
# _Outptr_ VOID **ppv);
PSCreatePropertyStoreFromObject = (
    propsys.PSCreatePropertyStoreFromObject
)
PSCreatePropertyStoreFromObject.restype = HRESULT

# punkSource must support IPropertyStore
# riid may be IPropertyStore, IPropertySetStorage,
# IPropertyStoreCapabilities, or IObjectProvider
# HRESULT PSCreateAdapterFromPropertyStore(
# _In_ IPropertyStore *pps,
# _In_ REFIID riid,
# _Outptr_ VOID **ppv);
PSCreateAdapterFromPropertyStore = (
    propsys.PSCreateAdapterFromPropertyStore
)
PSCreateAdapterFromPropertyStore.restype = HRESULT

# Talk to the property system using an interface
# Returns an IPropertySystem interface

# HRESULT PSGetPropertySystem(
# _In_ REFIID riid,
# _Outptr_ VOID **ppv);
PSGetPropertySystem = propsys.PSGetPropertySystem
PSGetPropertySystem.restype = HRESULT

# Obtain a value from serialized property storage
# HRESULT PSGetPropertyFromPropertyStorage(
# _In_reads_bytes_(cb) PCUSERIALIZEDPROPSTORAGE psps,
# _In_ DWORD cb,
# _In_ REFPROPERTYKEY rpkey,
# _Out_ PROPVARIANT *ppropvar);
PSGetPropertyFromPropertyStorage = (
    propsys.PSGetPropertyFromPropertyStorage
)
PSGetPropertyFromPropertyStorage.restype = HRESULT

# Obtain a named value from serialized property storage
# HRESULT PSGetNamedPropertyFromPropertyStorage(
# _In_reads_bytes_(cb) PCUSERIALIZEDPROPSTORAGE psps,
# _In_ DWORD cb,
# _In_ LPCWSTR pszName,
# _Out_ PROPVARIANT *ppropvar);
PSGetNamedPropertyFromPropertyStorage = (
    propsys.PSGetNamedPropertyFromPropertyStorage
)
PSGetNamedPropertyFromPropertyStorage.restype = HRESULT

# Helper functions for reading and writing values from IPropertyBag's.
# HRESULT PSPropertyBag_ReadType(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _Out_ VARIANT *var,
# VARTYPE type);
PSPropertyBag_ReadType = propsys.PSPropertyBag_ReadType
PSPropertyBag_ReadType.restype = HRESULT

# HRESULT PSPropertyBag_ReadStr(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _Out_writes_(CHARacterCount) LPWSTR value,
# INT characterCount);
PSPropertyBag_ReadStr = propsys.PSPropertyBag_ReadStr
PSPropertyBag_ReadStr.restype = HRESULT

# HRESULT PSPropertyBag_ReadStrAlloc(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _Outptr_ PWSTR *value);
PSPropertyBag_ReadStrAlloc = propsys.PSPropertyBag_ReadStrAlloc
PSPropertyBag_ReadStrAlloc.restype = HRESULT

# HRESULT PSPropertyBag_ReadBSTR(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _Outptr_ BSTR *value);
PSPropertyBag_ReadBSTR = propsys.PSPropertyBag_ReadBSTR
PSPropertyBag_ReadBSTR.restype = HRESULT

# HRESULT PSPropertyBag_WriteStr(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _In_ LPCWSTR value);
PSPropertyBag_WriteStr = propsys.PSPropertyBag_WriteStr
PSPropertyBag_WriteStr.restype = HRESULT

# HRESULT PSPropertyBag_WriteBSTR(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _In_ BSTR value);
PSPropertyBag_WriteBSTR = propsys.PSPropertyBag_WriteBSTR
PSPropertyBag_WriteBSTR.restype = HRESULT

# HRESULT PSPropertyBag_ReadInt(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _Out_ INT *value);
PSPropertyBag_ReadInt = propsys.PSPropertyBag_ReadInt
PSPropertyBag_ReadInt.restype = HRESULT

# HRESULT PSPropertyBag_WriteInt(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# INT value);
PSPropertyBag_WriteInt = propsys.PSPropertyBag_WriteInt
PSPropertyBag_WriteInt.restype = HRESULT

# HRESULT PSPropertyBag_ReadSHORT(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _Out_ SHORT *value);
PSPropertyBag_ReadSHORT = propsys.PSPropertyBag_ReadSHORT
PSPropertyBag_ReadSHORT.restype = HRESULT

# HRESULT PSPropertyBag_WriteSHORT(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# SHORT value);
PSPropertyBag_WriteSHORT = propsys.PSPropertyBag_WriteSHORT
PSPropertyBag_WriteSHORT.restype = HRESULT

# HRESULT PSPropertyBag_ReadLONG(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _Out_ LONG *value);
PSPropertyBag_ReadLONG = propsys.PSPropertyBag_ReadLONG
PSPropertyBag_ReadLONG.restype = HRESULT

# HRESULT PSPropertyBag_WriteLONG(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# LONG value);
PSPropertyBag_WriteLONG = propsys.PSPropertyBag_WriteLONG
PSPropertyBag_WriteLONG.restype = HRESULT

# HRESULT PSPropertyBag_ReadDWORD(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _Out_ DWORD *value);
PSPropertyBag_ReadDWORD = propsys.PSPropertyBag_ReadDWORD
PSPropertyBag_ReadDWORD.restype = HRESULT

# HRESULT PSPropertyBag_WriteDWORD(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# DWORD value);
PSPropertyBag_WriteDWORD = propsys.PSPropertyBag_WriteDWORD
PSPropertyBag_WriteDWORD.restype = HRESULT

# HRESULT PSPropertyBag_ReadBOOL(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _Out_ BOOL *value);
PSPropertyBag_ReadBOOL = propsys.PSPropertyBag_ReadBOOL
PSPropertyBag_ReadBOOL.restype = HRESULT

# HRESULT PSPropertyBag_WriteBOOL(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# BOOL value);
PSPropertyBag_WriteBOOL = propsys.PSPropertyBag_WriteBOOL
PSPropertyBag_WriteBOOL.restype = HRESULT

# HRESULT PSPropertyBag_ReadPOINTL(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _Out_ POINTL *value);
PSPropertyBag_ReadPOINTL = propsys.PSPropertyBag_ReadPOINTL
PSPropertyBag_ReadPOINTL.restype = HRESULT

# HRESULT PSPropertyBag_WritePOINTL(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _In_ POINTL *value);
PSPropertyBag_WritePOINTL = propsys.PSPropertyBag_WritePOINTL
PSPropertyBag_WritePOINTL.restype = HRESULT

# HRESULT PSPropertyBag_ReadPOINTS(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _Out_ POINTS *value);
PSPropertyBag_ReadPOINTS = propsys.PSPropertyBag_ReadPOINTS
PSPropertyBag_ReadPOINTS.restype = HRESULT

# HRESULT PSPropertyBag_WritePOINTS(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _In_ POINTS *value);
PSPropertyBag_WritePOINTS = propsys.PSPropertyBag_WritePOINTS
PSPropertyBag_WritePOINTS.restype = HRESULT

# HRESULT PSPropertyBag_ReadRECTL(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _Out_ RECTL *value);
PSPropertyBag_ReadRECTL = propsys.PSPropertyBag_ReadRECTL
PSPropertyBag_ReadRECTL.restype = HRESULT

# HRESULT PSPropertyBag_WriteRECTL(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _In_ RECTL *value);
PSPropertyBag_WriteRECTL = propsys.PSPropertyBag_WriteRECTL
PSPropertyBag_WriteRECTL.restype = HRESULT

# HRESULT PSPropertyBag_ReadStream(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _Outptr_ IStream **value);
PSPropertyBag_ReadStream = propsys.PSPropertyBag_ReadStream
PSPropertyBag_ReadStream.restype = HRESULT

# HRESULT PSPropertyBag_WriteStream(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _In_ IStream *value);
PSPropertyBag_WriteStream = propsys.PSPropertyBag_WriteStream
PSPropertyBag_WriteStream.restype = HRESULT

# HRESULT PSPropertyBag_Delete(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName);
PSPropertyBag_Delete = propsys.PSPropertyBag_Delete
PSPropertyBag_Delete.restype = HRESULT

# HRESULT PSPropertyBag_ReadULONGLONG(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _Out_ ULONGLONG *value);
PSPropertyBag_ReadULONGLONG = propsys.PSPropertyBag_ReadULONGLONG
PSPropertyBag_ReadULONGLONG.restype = HRESULT

# HRESULT PSPropertyBag_WriteULONGLONG(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# ULONGLONG value);
PSPropertyBag_WriteULONGLONG = propsys.PSPropertyBag_WriteULONGLONG
PSPropertyBag_WriteULONGLONG.restype = HRESULT

# HRESULT PSPropertyBag_ReadUnknown(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _In_ REFIID riid,
# _Outptr_ VOID **ppv);
PSPropertyBag_ReadUnknown = propsys.PSPropertyBag_ReadUnknown
PSPropertyBag_ReadUnknown.restype = HRESULT

# HRESULT PSPropertyBag_WriteUnknown(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _In_ IUnknown *punk);
PSPropertyBag_WriteUnknown = propsys.PSPropertyBag_WriteUnknown
PSPropertyBag_WriteUnknown.restype = HRESULT

# HRESULT PSPropertyBag_ReadGUID(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _Out_ GUID *value);
PSPropertyBag_ReadGUID = propsys.PSPropertyBag_ReadGUID
PSPropertyBag_ReadGUID.restype = HRESULT

# HRESULT PSPropertyBag_WriteGUID(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _In_ GUID *value);
PSPropertyBag_WriteGUID = propsys.PSPropertyBag_WriteGUID
PSPropertyBag_WriteGUID.restype = HRESULT

# HRESULT PSPropertyBag_ReadPropertyKey(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _Out_ PROPERTYKEY *value);
PSPropertyBag_ReadPropertyKey = propsys.PSPropertyBag_ReadPropertyKey
PSPropertyBag_ReadPropertyKey.restype = HRESULT

# HRESULT PSPropertyBag_WritePropertyKey(
# _In_ IPropertyBag *propBag,
# _In_ LPCWSTR propName,
# _In_ REFPROPERTYKEY value);
PSPropertyBag_WritePropertyKey = propsys.PSPropertyBag_WritePropertyKey
PSPropertyBag_WritePropertyKey.restype = HRESULT

# Additional Prototypes for ALL interfaces
oleaut32 = ctypes.windll.OLEAUT32

# ULONG BSTR_UserSize(
#     __RPC__in ULONG *, 
#     ULONG, 
#     __RPC__in BSTR *
# );
BSTR_UserSize = oleaut32.BSTR_UserSize
BSTR_UserSize.restype = ULONG

# UCHAR * BSTR_UserMarshal(
#     __RPC__in ULONG *, 
#     __RPC__inout_xcount(0) UCHAR *, 
#     __RPC__in BSTR * 
# );
BSTR_UserMarshal = oleaut32.BSTR_UserMarshal
BSTR_UserMarshal.restype = POINTER(UBYTE)

# UCHAR * BSTR_UserUnmarshal(
#     __RPC__in ULONG *, 
#     __RPC__in_xcount(0) UCHAR *,
#     __RPC__out BSTR * 
# );
BSTR_UserUnmarshal = oleaut32.BSTR_UserUnmarshal
BSTR_UserUnmarshal.restype = POINTER(UBYTE)

# VOID BSTR_UserFree(
#     __RPC__in ULONG *, 
#     __RPC__in BSTR * 
# );
BSTR_UserFree = oleaut32.BSTR_UserFree
BSTR_UserFree.restype = VOID

# ULONG LPSAFEARRAY_UserSize(
#     __RPC__in ULONG *, 
#     ULONG, 
#     __RPC__in LPSAFEARRAY *
# );
LPSAFEARRAY_UserSize = oleaut32.LPSAFEARRAY_UserSize
LPSAFEARRAY_UserSize.restype = ULONG

# UCHAR * LPSAFEARRAY_UserMarshal(
#     __RPC__in ULONG *, 
#     __RPC__inout_xcount(0) UCHAR *, 
#     __RPC__in LPSAFEARRAY * 
# );
LPSAFEARRAY_UserMarshal = oleaut32.LPSAFEARRAY_UserMarshal
LPSAFEARRAY_UserMarshal.restype = POINTER(UBYTE)

# UCHAR * LPSAFEARRAY_UserUnmarshal(
#     __RPC__in ULONG *, 
#     __RPC__in_xcount(0) UCHAR *, 
#     __RPC__out LPSAFEARRAY * 
# );
LPSAFEARRAY_UserUnmarshal = oleaut32.LPSAFEARRAY_UserUnmarshal
LPSAFEARRAY_UserUnmarshal.restype = POINTER(UBYTE)

# VOID LPSAFEARRAY_UserFree(
#     __RPC__in ULONG *, 
#     __RPC__in LPSAFEARRAY * 
# );
LPSAFEARRAY_UserFree = oleaut32.LPSAFEARRAY_UserFree
LPSAFEARRAY_UserFree.restype = VOID

# ULONG BSTR_UserSize64(
#     __RPC__in ULONG *, 
#     ULONG, 
#     __RPC__in BSTR * 
# );
BSTR_UserSize64 = oleaut32.BSTR_UserSize64
BSTR_UserSize64.restype = ULONG

# UCHAR * BSTR_UserMarshal64(
#     __RPC__in ULONG *, 
#     __RPC__inout_xcount(0) UCHAR *, 
#     __RPC__in BSTR *
# );
BSTR_UserMarshal64 = oleaut32.BSTR_UserMarshal64
BSTR_UserMarshal64.restype = POINTER(UBYTE)

# UCHAR * BSTR_UserUnmarshal64(
#     __RPC__in ULONG *, 
#     __RPC__in_xcount(0) UCHAR *, 
#     __RPC__out BSTR * 
# );
BSTR_UserUnmarshal64 = oleaut32.BSTR_UserUnmarshal64
BSTR_UserUnmarshal64.restype = POINTER(UBYTE)

# VOID BSTR_UserFree64(
#     __RPC__in ULONG *, 
#     __RPC__in BSTR * 
# );
BSTR_UserFree64 = oleaut32.BSTR_UserFree64
BSTR_UserFree64.restype = VOID

# ULONG LPSAFEARRAY_UserSize64(
#     __RPC__in ULONG *, 
#     ULONG, 
#     __RPC__in LPSAFEARRAY * 
# );
LPSAFEARRAY_UserSize64 = oleaut32.LPSAFEARRAY_UserSize64
LPSAFEARRAY_UserSize64.restype = ULONG

# UCHAR * LPSAFEARRAY_UserMarshal64(
#     __RPC__in ULONG *,
#     __RPC__inout_xcount(0) UCHAR *,
#     __RPC__in LPSAFEARRAY * 
# );
LPSAFEARRAY_UserMarshal64 = oleaut32.LPSAFEARRAY_UserMarshal64
LPSAFEARRAY_UserMarshal64.restype = POINTER(UBYTE)

# UCHAR * LPSAFEARRAY_UserUnmarshal64(
#     __RPC__in ULONG *,
#     __RPC__in_xcount(0) UCHAR *,
#     __RPC__out LPSAFEARRAY *
# );
LPSAFEARRAY_UserUnmarshal64 = oleaut32.LPSAFEARRAY_UserUnmarshal64
LPSAFEARRAY_UserUnmarshal64.restype = POINTER(UBYTE)

# VOID LPSAFEARRAY_UserFree64(
#     __RPC__in ULONG *, 
#     __RPC__in LPSAFEARRAY *
# );
LPSAFEARRAY_UserFree64 = oleaut32.LPSAFEARRAY_UserFree64
LPSAFEARRAY_UserFree64.restype = VOID
