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

import pyWinCoreAudio
import threading
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
    ON_PART_CHANGE
)


def on_device_added(signal, device):
    print('Device added:', device.name)


_on_device_added = ON_DEVICE_ADDED.register(on_device_added)


def on_device_removed(signal, name):
    print('Device removed:', name)


_on_device_removed = ON_DEVICE_REMOVED.register(on_device_removed)


def on_device_property_changed(signal, device, key, endpoint=None):
    if endpoint is not None:
        print('Property changed:', device.name + '.' + endpoint.name, key)

    else:
        print('Property changed:', device.name, key)


_on_device_property_changed = ON_DEVICE_PROPERTY_CHANGED.register(on_device_property_changed)


def on_device_state_changed(signal, device, new_state, endpoint=None):
    if endpoint is not None:
        print('State changed:', device.name + '.' + endpoint.name, new_state)

    else:
        print('State changed:', device.name, new_state)


_on_device_state_changed = ON_DEVICE_STATE_CHANGED.register(on_device_state_changed)


def on_endpoint_volume_changed(signal, device, endpoint, is_muted, master_volume,  channel_volumes):
    print('Endpoint volume changed:', device.name + '.' + endpoint.name)
    print('mute:', is_muted)
    print('volume:', master_volume)
    for i, channel in enumerate(channel_volumes):
        print('channel:', i)
        print('volume:', channel)


_on_endpoint_volume_changed = ON_ENDPOINT_VOLUME_CHANGED.register(on_endpoint_volume_changed)


def on_endpoint_default_changed(signal, device, endpoint, role, flow):
    print('Default changed:', device.name + '.' + endpoint.name, 'role:', role, 'flow:', flow)


_on_endpoint_default_changed = ON_ENDPOINT_DEFAULT_CHANGED.register(on_endpoint_default_changed)


def on_session_volume_duck(signal, device, endpoint, session, count_communication_sessions):
    print('Session volume duck:', device.name + '.' + endpoint.name + '.' + session.name, count_communication_sessions)


_on_session_volume_duck = ON_SESSION_VOLUME_DUCK.register(on_session_volume_duck)


def on_session_volume_unduck(signal, device, endpoint, session):
    print('Session volume unduck:', device.name + '.' + endpoint.name + '.' + session.name)


_on_session_volume_unduck = ON_SESSION_VOLUME_UNDUCK.register(on_session_volume_unduck)


def on_session_created(signal, device, endpoint, session):
    print('Session created:', device.name + '.' + endpoint.name + '.' + session.name)


_on_session_created = ON_SESSION_CREATED.register(on_session_created)


def on_session_name_changed(signal, device, endpoint, session, old_name, new_name):
    print('Session name changed:', device.name + '.' + endpoint.name + '.' + old_name, new_name)


_on_session_name_changed = ON_SESSION_NAME_CHANGED.register(on_session_name_changed)


def on_session_grouping_changed(signal, device, endpoint, session, new_grouping_param):
    print('Session grouping changed:', device.name + '.' + endpoint.name + '.' + session.name, new_grouping_param)


_on_session_grouping_changed = ON_SESSION_GROUPING_CHANGED.register(on_session_grouping_changed)


def on_session_icon_changed(signal, device, endpoint, session, old_icon, new_icon):
    print('Session icon changed:', device.name + '.' + endpoint.name + '.' + session.name, 'old:', old_icon, 'new:', new_icon)


_on_session_icon_changed = ON_SESSION_ICON_CHANGED.register(on_session_icon_changed)


def on_session_disconnect(signal, device, endpoint, name, reason):
    print('Session disconnected:', device.name + '.' + endpoint.name + '.' + name, reason)


_on_session_disconnect = ON_SESSION_DISCONNECT.register(on_session_disconnect)


def on_session_volume_changed(signal, device, endpoint, session, new_volume, new_mute):
    print('Session volume changed:', device.name + '.' + endpoint.name + '.' + session.name, 'volume:', new_volume, 'mute:', new_mute)


_on_session_volume_changed = ON_SESSION_VOLUME_CHANGED.register(on_session_volume_changed)


def on_session_state_changed(signal, device, endpoint, session, new_state):
    print('Session volume changed:', device.name + '.' + endpoint.name + '.' + session.name, 'state:', new_state)


_on_session_state_changed = ON_SESSION_STATE_CHANGED.register(on_session_state_changed)


def on_session_channel_volume_changed(signal, device, endpoint, session, channel, new_volume):
    print('Session channel volume changed:', device.name + '.' + endpoint.name + '.' + session.name, 'volume:', new_volume, 'channel:', channel)


_on_session_channel_volume_changed = ON_SESSION_CHANNEL_VOLUME_CHANGED.register(on_session_channel_volume_changed)


def on_part_changed(signal, device, endpoint, interface, process_id):
    print('Part changed:', device.name + '.' + endpoint.name, 'interface:', interface.__class__.__name__, 'process_id:', process_id)


_on_part_changed = ON_PART_CHANGE.register(on_part_changed)


devices = pyWinCoreAudio.devices()

for device in devices():
    print(device.name)
    print('    endpoints:')
    for endpoint in device:
        print('        endpoint:', endpoint.name)
        print('        description:', endpoint.description)
        print('        type:', endpoint.type)
        print('        data flow:', endpoint.data_flow)
        print('        form_factor:', endpoint.form_factor)
        print('        full_range_speakers:', endpoint.full_range_speakers)
        print('        guid:', endpoint.guid)
        print('        physical_speakers:', endpoint.physical_speakers)
        print('        system_effects:', endpoint.system_effects)
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
        print('        auto_gain_control:', endpoint.auto_gain_control)
        print('        bass:', endpoint.bass)
        print('        channel_config:', endpoint.channel_config)
        print('        input:', endpoint.input)
        print('        loudness:', endpoint.loudness)
        print('        midrange:', endpoint.midrange)
        print('        output:', endpoint.output)
        print('        treble:', endpoint.treble)
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
        del volume

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
            del volume
            del session

            print()

        del endpoint

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

        del connector

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

        del subunit

    del device


event = threading.Event()

try:
    event.wait()
except KeyboardInterrupt:
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

pyWinCoreAudio.stop()

