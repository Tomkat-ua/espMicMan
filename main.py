import aioesphomeapi
import asyncio
#import win32api
#import win32gui
import win32_mic_controller as micc

# APPCOMMAND_MIC_ON_OFF_TOGGLE      44 Toggle the microphone.
# APPCOMMAND_MICROPHONE_VOLUME_DOWN 25 Decrease microphone volume.
# APPCOMMAND_MICROPHONE_VOLUME_MUTE 24 Mute the microphone.
# APPCOMMAND_MICROPHONE_VOLUME_UP   26 Increase microphone volume.

host =  "192.168.2.32"
password =  "EspAdminn"

audio_control = micc.AudioUtilities()
async def mic_state():
    # WM_APPCOMMAND = 0x319
    # APPCOMMAND_MICROPHONE_VOLUME_MUTE = 0x180000 #24
    # APPCOMMAND_MIC_ON_OFF_TOGGLE      = 0x2C0000 #44
    # APPCOMMAND_MICROPHONE_VOLUME_DOWN = 0x190000 #25
    # APPCOMMAND_MICROPHONE_VOLUME_UP   = 0x1a0000 #26
    # hwnd_active = win32gui.GetForegroundWindow()
    # if state == True:
    #     win32api.SendMessage(hwnd_active, WM_APPCOMMAND, None, APPCOMMAND_MICROPHONE_VOLUME_UP)
    # if state == False:
    #     win32api.SendMessage(hwnd_active, WM_APPCOMMAND, None, APPCOMMAND_MICROPHONE_VOLUME_MUTE)
    #loop = asyncio.get_running_loop()
    audio_control = micc.AudioUtilities()
    print(audio_control.GetStateMute())

async def info():
    """Connect to an ESPHome device and get details."""
    loop = asyncio.get_running_loop()

    # Establish connection
    api = aioesphomeapi.APIClient(loop,host, 6053, password)
    await api.connect(login=True)

    # Get API version of the device's firmware
    print(api.api_version)

    # Show device details
    device_info = await api.device_info()
    print(device_info)

    # List all entities of the device
    entities = await api.list_entities_services()
    print(entities)

# loop = asyncio.get_event_loop()
# loop.run_until_complete(info())

async def main():

    """Connect to an ESPHome device and wait for state changes."""
    loop = asyncio.get_running_loop()
    cli = aioesphomeapi.APIClient(loop,host, 6053, password)

    await cli.connect(login=True)
    def change_callback(state):
        """Print the state changes of the device.."""
        #audio_control.MuteMicrophone()
        #audio_control.toggle_mute
        if str(state).find('key=3722332099') >=0:
            if str(state).find('state=True') >=0:
                print(state)
                audio_control.UnMuteMicrophone()
            if str(state).find('state=False') >=0:
                print(state)
                audio_control.MuteMicrophone()

    # Subscribe to the state changes
    await cli.subscribe_states(change_callback)

loop = asyncio.get_event_loop()

try:

    ##asyncio.ensure_future(info())
    #asyncio.run(mic_state())
    #asyncio.ensure_future(mic_state())
    asyncio.ensure_future(main())
    loop.run_forever()
except KeyboardInterrupt:
    print ('Exiting')
finally:
    loop.close()
