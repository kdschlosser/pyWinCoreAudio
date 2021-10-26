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

import weakref


class Singleton(type):
    _instances = {}

    @staticmethod
    def remove_instance(cls, instance):
        for key, value in list(cls._instances.items()):
            if instance == value:
                del cls._instances[key]
                break

    def __init__(cls, name, bases, dct):
        cls._instances = {}

        super(Singleton, cls).__init__(name, bases, dct)

    def __call__(cls, *args, **kwargs):
        key = list(args)

        for k in sorted(list(kwargs.keys())):
            key.append(kwargs[k])

        key = tuple(str(item) for item in key)

        if key in cls._instances:
            instance = cls._instances[key]()
            if instance is None:
                instance = super(
                    Singleton,
                    cls
                ).__call__(*args, **kwargs)

                cls._instances[key] = weakref.ref(instance)
        else:
            instance = super(
                Singleton,
                cls
            ).__call__(*args, **kwargs)

            cls._instances[key] = weakref.ref(instance)

        return instance





        return

        # print '{'
        #
        # for key, value in cls._instances.items():
        #     print '   ', repr(str(key)) + ': {'
        #     for k, v in value.items():
        #         k = ', '.join(repr(str(item)) for item in k)
        #         print '        (' + k + '):', repr(str(v)) + ','
        #     print '    },'
        # print '}'
