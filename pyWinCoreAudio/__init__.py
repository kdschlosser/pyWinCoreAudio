from _main import (
    AudioDevices as _AudioDevices,
    GetRGB
)
import comtypes


comtypes.CoInitialize()

AudioDevices = _AudioDevices()


def stop():
    comtypes.CoUninitialize()


if __name__ == '__main__':

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
                print
                print '    endpoint name:', endpoint.name
                print '    ---------------------------------------------------'
                print '        description:', endpoint.description
                print '        data flow:', endpoint.data_flow
                print '        form factor:', endpoint.form_factor
                print '        type:', endpoint.form_factor
                print '        full range speakers:', endpoint.full_range_speakers
                print '        guid:', endpoint.guid
                print '        physical speakers:', endpoint.physical_speakers
                print '        system effects:', endpoint.system_effects

                # try:
                #     sink = endpoint.jack_information
                # except AttributeError:
                #     pass
                # else:
                #     print '    manufacturer id:', sink.manufacturer_id
                #     print '    product id:', sink.product_id
                #     print '    audio latency:', sink.audio_latency
                #     print '    hdcp capable:', sink.hdcp_capable
                #     print '    ai capable:', sink.ai_capable
                #     print '    description:', sink.description
                #     print '    port id:', sink.port_id
                #     print '    connection type:', sink.connection_type

                try:
                    i = 0

                    for jack in endpoint.jack_descriptions:
                        print
                        print '        connector %d' % i
                        print '            channel mapping:', jack.channel_mapping
                        print '            color:', GetRGB(jack.color)
                        print '            type:', jack.type
                        print '            location:', jack.location
                        print '            port:', jack.port
                        print '            is connected:', jack.is_connected
                        i += 1
                except AttributeError:
                    pass

                try:
                    volume = endpoint.volume
                except AttributeError:
                    pass
                else:
                    print '        volume'
                    print '            master level:', volume.master
                    print '            master level scalar:', volume.master_scalar
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
                    print '            channel levels'
                    levels = volume.channel_levels
                    scalars = volume.channel_levels_scalar
                    ranges = volume.channel_ranges
                    for i, level in enumerate(levels):
                        print '            %d:' % i
                        print '                level:', level
                        print '                scalar level:', scalars[i]
                        c_min, c_max, c_step = ranges[i]
                        print '                min:', c_min
                        print '                max:', c_max
                        print '                step:', c_step
                    try:
                        peak_meter = volume.peak_meter
                        print '            peak meter'
                        peak_values = peak_meter.channel_peak_values
                        for i, value in enumerate(peak_values):
                            print '                %d' % i, value

                    except AttributeError:
                        pass


    try:
        test()
    finally:
        stop()
