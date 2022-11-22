
from .inspectable import IInspectable
from .data_types import *


IID_IActivationFactory = IID("{00000035-0000-0000-C000-000000000046}")


class IActivationFactory(IInspectable):
    _case_insensitive_ = False
    _iid_ = IID_IActivationFactory
    _methods_ = (
        comtypes.STDMETHOD(
            HRESULT,
            'ActivateInstance',
            [POINTER(POINTER(IInspectable))]
        ),
    )
