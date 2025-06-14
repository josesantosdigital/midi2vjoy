#  midi2vjoy.py
#
#  Copyright 2017  <c0redumb>

import sys, os, time, traceback
import ctypes
from optparse import OptionParser
import pygame.midi
import winreg

# Constants
# Axis mapping
axis = {'X': 0x30, 'Y': 0x31, 'Z': 0x32, 'RX': 0x33, 'RY': 0x34, 'RZ': 0x35,
        'SL0': 0x36, 'SL1': 0x37, 'WHL': 0x38, 'POV': 0x39}

# Sliders or Pitchbend MIDI types
sliders = {176, 224}

# Buttons with Note On/Off types
btns = {144, 128, 153, 137}

# Manual overrides if you want to treat a slider as a button
sliderOverride = {}

# Globals
options = None

def midi_test():
    n = pygame.midi.get_count()

    # List all the devices and make a choice
    print('Input MIDI devices:')
    for i in range(n):
        info = pygame.midi.get_device_info(i)
        if info[2]:
            print(i, info[1].decode())
    d = int(input('Select MIDI device to test: '))

    # Open the device for testing
    try:
        print('Opening MIDI device:', d)
        m = pygame.midi.Input(d)
        print('Device opened for testing. Use ctrl-c to quit.')
        while True:
            while m.poll():
                print(m.read(1))
            time.sleep(0.1)
    except:
        m.close()

def read_conf(conf_file):
    '''Read the configuration file'''
    table = {}
    vids = []
    with open(conf_file, 'r') as f:
        for l in f:
            if len(l.strip()) == 0 or l[0] == '#':
                continue
            fs = l.split()
            key = (int(fs[0]), int(fs[1]))
            if fs[0] == '144':
                val = (int(fs[2]), int(fs[3]))  # Note input -> vJoy button
            else:
                try:
                    val = (int(fs[2]), int(fs[3]))  # Treat as button (CC override)
                except ValueError:
                    val = (int(fs[2]), fs[3])  # Axis mapping
            table[key] = val
            vid = int(fs[2])
            if not vid in vids:
                vids.append(vid)
    return (table, vids)

def joystick_run():
    # Process the configuration file
    if options.conf is None:
        print('Must specify a configuration file')
        return
    try:
        if options.verbose:
            print('Opening configuration file:', options.conf)
        (table, vids) = read_conf(options.conf)
    except:
        print('Error processing the configuration file:', options.conf)
        return

    # Getting the MIDI device ready
    if options.midi is None:
        print('Must specify a MIDI interface to use')
        return
    try:
        if options.verbose:
            print('Opening MIDI device:', options.midi)
        midi = pygame.midi.Input(options.midi)
    except:
        print('Error opening MIDI device:', options.midi)
        return

    # Load vJoysticks
    try:
        vjoyregkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{8E31F76F-74C3-47F1-9550-E041EEDC5FBB}_is1')
        installpath = winreg.QueryValueEx(vjoyregkey, 'InstallLocation')
        winreg.CloseKey(vjoyregkey)
        dll_file = os.path.join(installpath[0], 'x64', 'vJoyInterface.dll')
        vjoy = ctypes.WinDLL(dll_file)

        for vid in vids:
            if options.verbose:
                print('Acquiring vJoystick:', vid)
            assert(vjoy.AcquireVJD(vid) == 1)
            assert(vjoy.GetVJDStatus(vid) == 0)
            vjoy.ResetVJD(vid)
    except:
        print('Error initializing virtual joysticks')
        return

    try:
        if options.verbose:
            print('Ready. Use ctrl-c to quit.')
        while True:
            while midi.poll():
                ipt = midi.read(1)
                key = tuple(ipt[0][0][0:2])
                reading = ipt[0][0][2]
                if not key in table:
                    continue
                opt = table[key]
                if options.verbose:
                    print(key, '->', opt, reading)

                if isinstance(opt[1], str) and opt[1] in axis:
                    # A slider/axis input (mapped to axis like Z, RX, etc.)
                    reading = (reading + 1) << 8
                    vjoy.SetAxis(reading, opt[0], axis[opt[1]])
                else:
                    # A button input (including CC-as-button)
                    # Treat value >= 64 as press (1), else release (0)
                    is_pressed = 1 if reading >= 64 else 0
                    vjoy.SetBtn(is_pressed, opt[0], int(opt[1]))
            time.sleep(0.01)
    except KeyboardInterrupt:
        pass
    except Exception as e:
        traceback.print_exc()

    # Relinquish vJoysticks
    for vid in vids:
        if options.verbose:
            print('Relinquishing vJoystick:', vid)
        vjoy.RelinquishVJD(vid)

    # Close MIDI device
    if options.verbose:
        print('Closing MIDI device')
    midi.close()

def main():
    # parse arguments
    parser = OptionParser()
    parser.add_option("-t", "--test", dest="runtest", action="store_true",
                      help="To test the midi inputs")
    parser.add_option("-m", "--midi", dest="midi", action="store", type="int",
                      help="ID of the MIDI input device to use")
    parser.add_option("-c", "--conf", dest="conf", action="store",
                      help="Configuration file for the translation")
    parser.add_option("-v", "--verbose", action="store_true", dest="verbose")
    parser.add_option("-q", "--quiet", action="store_false", dest="verbose")

    global options
    (options, args) = parser.parse_args()

    pygame.midi.init()

    if options.runtest:
        midi_test()
    else:
        joystick_run()

    pygame.midi.quit()

if __name__ == '__main__':
    main()
