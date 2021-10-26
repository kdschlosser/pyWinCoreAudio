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

from .__core_audio.constant import (
    SPEAKER_FRONT_LEFT,
    SPEAKER_FRONT_RIGHT,
    SPEAKER_FRONT_CENTER,
    SPEAKER_LOW_FREQUENCY,
    SPEAKER_BACK_LEFT,
    SPEAKER_BACK_RIGHT,
    SPEAKER_FRONT_LEFT_OF_CENTER,
    SPEAKER_FRONT_RIGHT_OF_CENTER,
    SPEAKER_BACK_CENTER,
    SPEAKER_SIDE_LEFT,
    SPEAKER_SIDE_RIGHT,
    SPEAKER_TOP_CENTER,
    SPEAKER_TOP_FRONT_LEFT,
    SPEAKER_TOP_FRONT_CENTER,
    SPEAKER_TOP_FRONT_RIGHT,
    SPEAKER_TOP_BACK_LEFT,
    SPEAKER_TOP_BACK_CENTER,
    SPEAKER_TOP_BACK_RIGHT
)


class AudioSpeakers(object):

    def __init__(self, value):
        if value is None:
            value = 0

        self.front_left = value | SPEAKER_FRONT_LEFT == value
        self.front_left_of_center = value | SPEAKER_FRONT_LEFT_OF_CENTER == value
        self.front_center = value | SPEAKER_FRONT_CENTER == value
        self.front_right_of_center = value | SPEAKER_FRONT_RIGHT_OF_CENTER == value
        self.front_right = value | SPEAKER_FRONT_RIGHT == value
        self.side_left = value | SPEAKER_SIDE_LEFT == value
        self.side_right = value | SPEAKER_SIDE_RIGHT == value
        self.back_left = value | SPEAKER_BACK_LEFT == value
        self.back_center = value | SPEAKER_BACK_CENTER == value
        self.back_right = value | SPEAKER_BACK_RIGHT == value
        self.high_center = value | SPEAKER_TOP_CENTER == value
        self.high_front_left = value | SPEAKER_TOP_FRONT_LEFT == value
        self.high_front_center = value | SPEAKER_TOP_FRONT_CENTER == value
        self.high_front_right = value | SPEAKER_TOP_FRONT_RIGHT == value
        self.high_back_left = value | SPEAKER_TOP_BACK_LEFT == value
        self.high_back_center = value | SPEAKER_TOP_BACK_CENTER == value
        self.high_back_right = value | SPEAKER_TOP_BACK_RIGHT == value
        self.subwoofer = value | SPEAKER_LOW_FREQUENCY == value

    def __str__(self):
        eye_level = sum([
            self.front_left,
            self.front_left_of_center,
            self.front_center,
            self.front_right_of_center,
            self.front_right,
            self.side_left,
            self.side_right,
            self.back_left,
            self.back_center,
            self.back_right
        ])

        three_d = sum([
            self.high_center,
            self.high_front_left,
            self.high_front_center,
            self.high_front_right,
            self.high_back_left,
            self.high_back_center,
            self.high_back_right
        ])

        sw = int(self.subwoofer)
        if three_d:
            return '{0}.{1}.{2}'.format(eye_level, sw, three_d)
        if eye_level:
            return '{0}.{1}'.format(eye_level, sw)
        if sw:
            return '{0}.{1}'.format(eye_level, sw)
        return '0'
