# Home Automations
# Call Indicator Light`

## This script updates the color and brightness of smart lights based on a video call status, or status of other 
## designated processes. It is designed to work with Home Assistant and smart lights that support color changes via 
## the Home Assistant REST API. This script monitors for video call applications (like Zoom) running on the local 
## machine. When a video call application is detected, it changes the color and brightness of specified smart lights to
## an alert color (e.g., bright pink) to indicate that the user is in a video call. When the video call application is no
## longer detected, it restores the lights to their previous state. Configuration settings are read from a JSON file. The
## script  runs in an infinite loop, checking the state of the specified processes every 30 seconds. Note: Ensure that 
## the settings.json file contains the correct Home Assistant base paths, token, process names, entity IDs, and 
## color/brightness settings.

## Usually, it is best to run this script as a scheduled task at logon. I find lightbulbs that are located outside of 
## the office let others know you are in a call. Make sure the lights you use are not used for other purposes, as this 
## script will change their color and brightness when a call is detected. If a light is on a non-alert color when a call 
## starts, it will change to the alert color and brightness, and then revert back to the previous color and brightness 
## when the call ends, or off if the light was off. If the light is in the alert color when the script starts, or when 
## not in a call, it will set the light to a default color and brightness specified in the JSON configuration file.
