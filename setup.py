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


try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

import pyWinCoreAudio


__license__ = open('LICENSE.md').read()
__doc__ = open('README.md').read()

setup(
    name='pyWinCoreAudio',
    version=pyWinCoreAudio.__version__,
    author=pyWinCoreAudio.__author__,
    description=pyWinCoreAudio.__description__,
    url=pyWinCoreAudio.__url__,
    license=__license__,
    keywords=['Windows Core Audio', 'Windows Audio', 'MMDeviceAPI', 'DeviceTopologyAPI'],
    packages=['pyWinCoreAudio'],
    long_description=__doc__,
    include_package_data=True,
    install_requires=['comtypes>=1.1.10'],
    zip_safe=True,
    classifiers=[
        'Intended Audience :: Developers',
        'License :: OSI Approved :: GNU General Public License v2 or later (GPLv2+)',
        'Natural Language :: English',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Development Status :: 3 - Alpha',
        'Environment :: Win32 (MS Windows)',
        'Operating System :: Microsoft :: Windows',
        'Operating System :: Microsoft :: Windows :: Windows 7',
        'Operating System :: Microsoft :: Windows :: Windows 8',
        'Operating System :: Microsoft :: Windows :: Windows 8.1',
        'Operating System :: Microsoft :: Windows :: Windows 10',
        'Operating System :: Microsoft :: Windows :: Windows 11',
        'Operating System :: Microsoft :: Windows :: Windows Vista',
        'Operating System :: Microsoft :: Windows :: Windows Server 2008',
        'Topic :: Multimedia :: Sound/Audio'
    ],
)
