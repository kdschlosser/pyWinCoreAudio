from _main import (
    AudioDevices as _AudioDevices,
)
import comtypes


comtypes.CoInitialize()

AudioDevices = _AudioDevices()


def stop():
    comtypes.CoUninitialize()


if __name__ == '__main__':
    registered_volume_callbacks = {}


    def volume_change(endpoint, master_volume, channel_volumes, mute):
        print 'volume event'
        print '    device name:', endpoint.device.name
        print '    endpoint name:', endpoint.name
        print '    master:', str(master_volume * 100) + '%'
        print '    mute:', mute
        for i, level in enumerate(channel_volumes):
            print '    channel %d level:' % i, str(level * 100) + '%'


    def state_change(device, state):
        print 'state change'
        print '    device name:', device.name
        print '    state:', state


    def device_added(device):
        print 'device added'
        print '    device name:', device.name


    def device_removed(device):
        print 'device removed'
        print '    device name:', device.name


    def default_changed(device):
        print 'default changed'
        print '    device name:', device.name

        for endpoint, callback in registered_volume_callbacks.items()[:]:
            if not endpoint.is_default:
                endpoint.volume.unregister_volume_change_callback(callback)
                del registered_volume_callbacks[endpoint]
                break

        for endpoint in device:
            if (
                endpoint.is_default and
                endpoint not in registered_volume_callbacks
            ):
                print '    endpoint name:', endpoint.name
                try:
                    callback = endpoint.volume.register_volume_change_callback(
                        volume_change
                    )
                    registered_volume_callbacks[endpoint] = callback
                    break

                except AttributeError:
                    pass


    def property_changed(device, key):
        print 'property changed'
        print '    device name:', device.name
        print '    key:', key.fmtid, key.pid

    def test():
        for device in AudioDevices:
            print
            print
            print
            print 'device name:', device.name
            print '======================================================='
            print '    id:', device.id
            print '    connector count:', device.connector_count
            print '    state:', device.state
            for endpoint in device:
                full_range_speakers = endpoint.full_range_speakers
                print
                print '    endpoint name:', endpoint.name
                print '    ---------------------------------------------------'
                print '        description:', endpoint.description
                print '        data flow:', endpoint.data_flow
                print '        form factor:', endpoint.form_factor
                print '        type:', endpoint.form_factor
                print '        full range speakers:', full_range_speakers
                print '        guid:', endpoint.guid
                print '        physical speakers:', endpoint.physical_speakers
                print '        system effects:', endpoint.system_effects

                try:
                    sink = endpoint.jack_information
                except AttributeError:
                    pass
                else:
                    print '        manufacturer id:', sink.manufacturer_id
                    print '        product id:', sink.product_id
                    print '        audio latency:', sink.audio_latency
                    print '        hdcp capable:', sink.hdcp_capable
                    print '        ai capable:', sink.ai_capable
                    print '        description:', sink.description
                    print '        connection type:', sink.connection_type

                try:
                    jacks = list(jack for jack in endpoint.jack_descriptions)
                    print
                    print '        connectors'

                    def p(attr, v):
                        print '               ', attr, v

                    for i, jack in enumerate(jacks):
                        print '            %d:' % i
                        p('channel mapping:', jack.channel_mapping)
                        p('color:', jack.color)
                        p('type:', jack.type)
                        p('location:', jack.location)
                        p('port:', jack.port)
                        p('presence detection:', jack.presence_detection)
                        p('dynamic format change:', jack.dynamic_format_change)
                        p('is connected:', jack.is_connected)
                except AttributeError:
                    pass

                try:
                    volume = endpoint.volume
                    print
                except AttributeError:
                    pass
                else:
                    scalar = volume.master_scalar
                    print '        volume'
                    print '            master level:', volume.master
                    print '            master level scalar:', scalar
                    m_min, m_max, m_step = volume.range
                    print '            master min:', m_min
                    print '            master max:', m_max
                    print '            master step:', m_step
                    try:
                        mute = volume.mute
                        print '            mute:', mute
                    except AttributeError:
                        pass
                    print '            channel count:', volume.channel_count
                    print
                    print '            channel levels'
                    levels = volume.channel_levels
                    scalars = volume.channel_levels_scalar
                    ranges = volume.channel_ranges
                    for i, level in enumerate(levels):
                        print '                %d:' % i
                        print '                    level:', level
                        print '                    scalar level:', scalars[i]
                        c_min, c_max, c_step = ranges[i]
                        print '                    min:', c_min
                        print '                    max:', c_max
                        print '                    step:', c_step
                    try:
                        peak_meter = volume.peak_meter
                        print
                        print '            peak meter'
                        print '                master:', peak_meter.peak_value
                        print '                channels'
                        peak_values = peak_meter.channel_peak_values
                        for i, value in enumerate(peak_values):
                            print '                    %d:' % i
                            print '                        level:', value

                    except AttributeError:
                        pass

        print
        print
        default_render = AudioDevices.default_render_endpoint
        print 'default render device:', default_render.device.name
        print 'default render endpoint:', default_render.name
        print 'default render endpoint volume:', str(
            default_render.volume.master_scalar * 100
        ) + '%'

        default_capture = AudioDevices.default_capture_endpoint
        print 'default capture device:', default_capture.device.name
        print 'default capture endpoint:', default_capture.name
        print 'default capture endpoint volume:', str(
            default_capture.volume.master_scalar * 100
        ) + '%'

        registered_volume_callbacks[default_render] = (
            default_render.volume.register_volume_change_callback(
                volume_change
            )
        )
        registered_volume_callbacks[default_capture] = (
            default_capture.volume.register_volume_change_callback(
                volume_change
            )
        )

        AudioDevices.bind('property', property_changed)
        AudioDevices.bind('state', state_change)
        AudioDevices.bind('default', default_changed)
        AudioDevices.bind('remove', device_removed)
        AudioDevices.bind('add', device_added)

        print
        print

        raw_input(
            'Change the volume.\n'
            'Set the mute.\n'
            'Change the default device.\n'
            'Change some of the device properties.\n'
            '\n'
            'Press any key to exit.\n'
        )
        import sys

        sys.exit()

    try:
        test()
    finally:
        stop()
