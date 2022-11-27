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


import sys
import os


sys.path.insert(0, os.path.dirname(__file__))

import pyWinCoreAudio
from pyWinCoreAudio import (
    ON_DEVICE_ADDED,
    ON_DEVICE_REMOVED,
    ON_DEVICE_PROPERTY_CHANGED,
    ON_DEVICE_STATE_CHANGED,
    ON_ENDPOINT_VOLUME_CHANGED,
    ON_ENDPOINT_DEFAULT_CHANGED,
    ON_SESSION_VOLUME_DUCK,
    ON_SESSION_VOLUME_UNDUCK,
    ON_SESSION_CREATED,
    ON_SESSION_NAME_CHANGED,
    ON_SESSION_GROUPING_CHANGED,
    ON_SESSION_ICON_CHANGED,
    ON_SESSION_DISCONNECT,
    ON_SESSION_VOLUME_CHANGED,
    ON_SESSION_STATE_CHANGED,
    ON_SESSION_CHANNEL_VOLUME_CHANGED,
    ON_PART_CHANGE,
    PKEY_AudioEndpoint_Disable_SysFx
)


def on_device_added(signal, device_):
    print(signal, ':', device_.name)


_on_device_added = ON_DEVICE_ADDED.register(on_device_added)


def on_device_removed(signal, name):
    print(signal, ':', name)


_on_device_removed = ON_DEVICE_REMOVED.register(on_device_removed)


def on_device_property_changed(signal, device_, key, endpoint_=None):
    if endpoint_ is not None:
        # if key in (
        #     pyWinCoreAudio.PKEY_AudioEndpoint_FullRangeSpeakers,
        #     pyWinCoreAudio.PKEY_AudioEndpoint_RealtekChannelsState,
        #     pyWinCoreAudio.PKEY_AudioEndpoint_RealtekChannelLevelOffset,
        #     pyWinCoreAudio.PKEY_AudioEndpoint_RealtekChannelDistance,
        # ):
        #     for speaker in endpoint.speakers:
        #         try:
        #             print('speaker id:', speaker.id)
        #             print('speaker active:', speaker.active)
        #             print('speaker full_range:', speaker.full_range)
        #             print('speaker level_offset:', speaker.level_offset)
        #             print('speaker distance:', speaker.distance)
        #         except:
        #             pass

        if key == PKEY_AudioEndpoint_Disable_SysFx:
            value = not endpoint_.audio_enhancements_enabled
            print(
                signal, ':',
                device_.name + '.' + endpoint_.name,
                key, '=', value
            )

        elif key in (pyWinCoreAudio.PKEY_AudioEndpoint_RealtekChannelDistance, pyWinCoreAudio.PKEY_AudioEndpoint_RealtekChannelLevelOffset):
            import winreg
            import struct

            path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{0}\{1}\FxProperties'.format(str(endpoint_.data_flow), endpoint_.guid.lower())

            ky = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_READ)
            fmtid = key.fmtid
            pid = key.pid
            try:
                value, d_type = winreg.QueryValueEx(ky, "{fmtid},{pid}".format(fmtid=str(fmtid).lower(), pid=pid))
            except:
                winreg.CloseKey(ky)
            else:
                winreg.CloseKey(ky)
                unpacked = list(item[0] for item in struct.iter_unpack('<h', value))
                print(key, "{fmtid},{pid}".format(fmtid=str(fmtid).lower(), pid=pid), unpacked)
                unpacked = list(item[0] for item in struct.iter_unpack('<H', value))
                print(unpacked)
                unpacked = list(item[0] for item in struct.iter_unpack('<Q', value))
                print(unpacked[0] & 0xFFFFFF, unpacked[0] >> 48)

        else:
            value = endpoint_.get_property(key, propvar=True)
            print(key.fmtid, ':', key.pid)
            print(
                signal, ':',
                device_.name + '.' + endpoint_.name,
                key, '=', value.vt, ':', value.value
            )

    else:
        print(signal, ':', device_.name, key)


    #                        FL    FR    FC          BL    BR    SL    SR
    # [65, 62143, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]

    #                  FL    FR    FC    LFE   BL    BR    SL    SR
    # [65, 102  , 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
# 1111 1111 1111 1111 1111 1111 1111 1111 1111 1111 1111 1111 1111 0010 1011 1111
# FF FF FF FF FF FF F2 BF
_on_device_property_changed = (
    ON_DEVICE_PROPERTY_CHANGED.register(on_device_property_changed)
)


def on_device_state_changed(signal, device_, new_state, endpoint_=None):
    if endpoint_ is not None:
        print(signal, ':', device_.name + '.' + endpoint_.name, new_state)
    else:
        print(signal, ':', device_.name, new_state)


_on_device_state_changed = (
    ON_DEVICE_STATE_CHANGED.register(on_device_state_changed)
)


def on_endpoint_volume_changed(
        signal,
        device_,
        endpoint_,
        is_muted,
        master_volume,
        channel_volumes
):
    print(signal, ':', device_.name + '.' + endpoint_.name)
    print('mute:', is_muted)
    print('volume:', master_volume)
    for i, channel_ in enumerate(channel_volumes):
        print('channel:', i)
        print('volume:', channel_)


_on_endpoint_volume_changed = (
    ON_ENDPOINT_VOLUME_CHANGED.register(on_endpoint_volume_changed)
)


def on_endpoint_default_changed(signal, device_, endpoint_, role, flow):
    print(
        signal, ':',
        device_.name + '.' + endpoint_.name,
        'role:', role,
        'flow:', flow
    )


_on_endpoint_default_changed = (
    ON_ENDPOINT_DEFAULT_CHANGED.register(on_endpoint_default_changed)
)


def on_session_volume_duck(
        signal,
        device_,
        endpoint_,
        session_,
        count_communication_sessions
):
    print(
        signal, ':',
        device_.name + '.' + endpoint_.name + '.' + session_.name,
        count_communication_sessions
    )


_on_session_volume_duck = (
    ON_SESSION_VOLUME_DUCK.register(on_session_volume_duck)
)


def on_session_volume_unduck(signal, device_, endpoint_, session_):
    print(
        signal, ':',
        device_.name + '.' + endpoint_.name + '.' + session_.name
    )


_on_session_volume_unduck = (
    ON_SESSION_VOLUME_UNDUCK.register(on_session_volume_unduck)
)


def on_session_created(signal, device_, endpoint_, session_):
    print(
        signal, ':',
        device_.name + '.' + endpoint_.name + '.' + session_.name
    )


_on_session_created = ON_SESSION_CREATED.register(on_session_created)


def on_session_name_changed(
        signal,
        device_,
        endpoint_,
        __,  # session_,
        old_name,
        new_name
):
    print(
        signal, ':',
        device_.name + '.' + endpoint_.name + '.' + old_name,
        device_.name + '.' + endpoint_.name + '.' + new_name
    )


_on_session_name_changed = (
    ON_SESSION_NAME_CHANGED.register(on_session_name_changed)
)


def on_session_grouping_changed(
        signal,
        device_,
        endpoint_,
        session_,
        new_grouping_param
):
    print(
        signal, ':',
        device_.name + '.' + endpoint_.name + '.' + session_.name,
        new_grouping_param
    )


_on_session_grouping_changed = (
    ON_SESSION_GROUPING_CHANGED.register(on_session_grouping_changed)
)


def on_session_icon_changed(
        signal,
        device_,
        endpoint_,
        session_,
        old_icon,
        new_icon
):
    print(
        signal, ':',
        device_.name + '.' + endpoint_.name + '.' + session_.name,
        'old:', old_icon,
        'new:', new_icon
    )


_on_session_icon_changed = (
    ON_SESSION_ICON_CHANGED.register(on_session_icon_changed)
)


def on_session_disconnect(signal, device_, endpoint_, name, reason):
    print(
        signal, ':',
        device_.name + '.' + endpoint_.name + '.' + name,
        reason
    )


_on_session_disconnect = ON_SESSION_DISCONNECT.register(on_session_disconnect)


def on_session_volume_changed(
        signal,
        device_,
        endpoint_,
        session_,
        new_volume,
        new_mute
):
    print(
        signal, ':',
        device_.name + '.' + endpoint_.name + '.' + session_.name,
        'volume:', new_volume,
        'mute:', new_mute
    )

    #
    # if new_volume <= 15.0:
    #     print('setting session volume', session.volume.level)
    #     session.volume.level = 50.0
    #     print('new session volume =', session.volume.level)


_on_session_volume_changed = (
    ON_SESSION_VOLUME_CHANGED.register(on_session_volume_changed)
)


def on_session_state_changed(
        signal,
        device_,
        endpoint_,
        session_,
        new_state
):
    print(
        signal, ':',
        device_.name + '.' + endpoint_.name + '.' + session_.name,
        'state:', new_state
    )


_on_session_state_changed = (
    ON_SESSION_STATE_CHANGED.register(on_session_state_changed)
)


def on_session_channel_volume_changed(
        signal,
        device_,
        endpoint_,
        session_,
        channel_,
        new_volume
):
    print(
        signal, ':',
        device_.name + '.' + endpoint_.name + '.' + session_.name,
        'volume:', new_volume,
        'channel:', channel_
    )


_on_session_channel_volume_changed = (
    ON_SESSION_CHANNEL_VOLUME_CHANGED.register(
        on_session_channel_volume_changed
    )
)


def on_part_changed(signal, device_, endpoint_, interface_, process_id):
    print(
        signal, ':',
        device_.name + '.' + endpoint_.name,
        'interface:', interface_.__class__.__name__,
        'process_id:', process_id
    )


_on_part_changed = ON_PART_CHANGE.register(on_part_changed)


for device in pyWinCoreAudio:
    print(device.name)
    print('    endpoints:')
    for endpoint in device:
        # if 'SPDIF' in endpoint.name:
        #     spdif_endpoint = endpoint

        print('        endpoint:', endpoint.name)
        print('        description:', endpoint.description)
        print('        type:', endpoint.type)
        print('        data flow:', endpoint.data_flow)
        print('        form_factor:', endpoint.form_factor)
        print('        guid:', endpoint.guid)
        print('        audio_enhancements_enabled:', endpoint.audio_enhancements_enabled)
        print('        hdcp_capable:', endpoint.hdcp_capable)
        print('        ai_capable:', endpoint.ai_capable)
        print('        connector_type:', endpoint.connector_type)
        print('        connector_name:', endpoint.connector_name)
        print('        connector_subtype:', endpoint.connector_subtype)
        print('        connector_location:', endpoint.connector_location)
        print('        connector_style:', endpoint.connector_style)
        print('        presence_detection:', endpoint.presence_detection)
        print('        connector_color:', endpoint.connector_color)
        print('        is_connected:', endpoint.is_connected)
        print('        channel_config:', endpoint.channel_config)
        print('        input:', endpoint.input)
        print('        output:', endpoint.output)
        print('        audio_channels:')

        # Chorus
        # modulation_rate
        # modulation_depth
        #
        # Reverb
        # reverb_time
        # delay_feedback

        for channel in endpoint.audio_channels:
            print('            channel num:', channel.channel_num)
            print('            bass bost:', channel.bass_boost)
            print('            loudness:', channel.loudness)
            print('            automatic gain control:', channel.automatic_gain_control)
            print('            bass:', channel.bass)
            print('            mid:', channel.mid)
            print('            treble:', channel.treble)
            print('            channel eq:')
            for band in channel.eq.bands:
                print('                frequency:', band.frequency)
                print('                level: ', band.level)
                print()
            print()
        print()

        print('        speakers:', endpoint.speakers)

        for speaker in endpoint.speakers:
            print('           ', speaker.name)
            print('                active:', speaker.active)
            print('                distance:', speaker.distance)
            print('                fullrange:', speaker.full_range)
            print('                level offset:', speaker.level_offset)
            print()

        volume = endpoint.volume
        if volume is not None:

            print('        volume:', volume.level)
            print('        volume db:', volume.level_db)
            print('        volume db min:', volume.db_minimum)
            print('        volume db max:', volume.db_maximum)
            print('        volume db increment:', volume.db_increment)
            print('        volume peak:', volume.peak)
            print('        volume mute:', volume.mute)
            print()

            for channel in volume:
                print('        channel:', channel.channel_number)
                print('        volume:', channel.level)
                print('        volume db:', channel.level_db)
                print('        volume db min:', channel.db_minimum)
                print('        volume db max:', channel.db_maximum)
                print('        volume db increment:', channel.db_increment)
                print('        volume peak:', channel.peak)
                print()

        print()
        print()

        print('        sessions:')

        for session in endpoint:
            print('            name:', session.name)
            print('            id:', session.id)
            print('            instance_id:', session.instance_id)
            print('            process_id:', session.process_id)
            print('            state:', session.state)
            print('            icon_path:', session.icon_path)
            print('            grouping_param:', session.grouping_param)
            print('            is_system_sounds:', session.is_system_sounds)
            volume = session.volume
            if volume is not None:
                print('            volume:', volume.level)
                print('            volume mute:', volume.mute)
                print()

                for channel in volume:
                    print('            channel:', channel.channel_number)
                    print('            volume:', channel.level)
                    print()

            print()
        print('        properties:')

        for k in endpoint.property_keys:
            value = k.get(endpoint)
            print('            key:', k)
            print('            value:', value)
            print()


            # If you want to change the session (application) endpoint
            # this is how you would go about doing it

            # if session.name == 'Mozilla Firefox':
            #     for edpt in device:
            #         if 'SPDIF' in edpt.name:
            #             session.default_endpoint = edpt
            #             print('session.default_endpoint == edpt:', session.default_endpoint == edpt)
            #
            #         del edpt

    print('    connectors:')

    for connector in device.connectors:
        part = connector.part
        print('        name:', part.name)
        print('        data flow:', connector.data_flow)
        print('        type:', connector.type)
        print('        connected:', connector.is_connected)
        print('        connector type:', part.type)
        print('        connector sub type:', part.sub_type)
        print('        interfaces:')
        for interface in part:
            print('           ', interface.name)
        print()

    print('    subunits:')
    for subunit in device.subunits:
        part = subunit.part
        print('        name:', part.name)
        print('        type:', part.type)
        print('        sub type:', part.sub_type)
        print('        interfaces:')
        for interface in part:
            print('           ', interface.name)
        print()


import signal as sig
import time


class ServiceExit(Exception):
    """
    Custom exception which is used to trigger the clean exit
    of all running threads and the main program.
    """
    pass


def service_shutdown(signum, frame):
    print('Caught signal %d' % signum)
    raise ServiceExit


sig.signal(sig.SIGTERM, service_shutdown)
sig.signal(sig.SIGINT, service_shutdown)

try:
    while True:
        time.sleep(0.5)

except ServiceExit:
    pass

_on_device_added.unregister()
_on_device_removed.unregister()
_on_device_property_changed.unregister()
_on_device_state_changed.unregister()
_on_endpoint_volume_changed.unregister()
_on_endpoint_default_changed.unregister()
_on_session_volume_duck.unregister()
_on_session_volume_unduck.unregister()
_on_session_created.unregister()
_on_session_name_changed.unregister()
_on_session_grouping_changed.unregister()
_on_session_icon_changed.unregister()
_on_session_disconnect.unregister()
_on_session_volume_changed.unregister()
_on_session_state_changed.unregister()
_on_session_channel_volume_changed.unregister()
_on_part_changed.unregister()

pyWinCoreAudio.unload()

