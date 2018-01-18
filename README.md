pyWinCoreAudio
==============

This is a work in progress. If there is anyone willing to help out.
Please by all means submit some PR's

This is a library that is going to remove the complexity
of controlling Windows audio devices. It is going to allow you to
change/view the following.

* Volume Level
* Channel Volume Level
* Mute
* Device Name
* Device description
* Endpoint Description
* Number of Endpoints
* Number of connectors
* Connector Type
* Render Devices
* Capture Devices
* Default Device
* Input
* Output
* Setting Loopback


for the following you can register for event callbacks.

* Audio Playing/Stopping Events
* Endpoint Changes
* Volume Changes
* Device Removal
* Device Addition


The following items can be changed/viewed if your sound card supports
it.

* Bass
* Treble
* Midrange
* Auto Gain Control
* Loudness
* Peak Meter
* Speaker Position
* Prologic Encoder/Decoder
* DNR (Dynamic Range Compression)
* Up/Down Mix
* Parametric Equalizer
* Equalizer
* 3D Effects
* DSP (Dynamic Sound Processing)
* Audio Delay
* Noise Suppression


My Goal is to create a single package for all of the
Windows Core Audio API.









