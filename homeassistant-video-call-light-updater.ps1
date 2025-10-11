# Program Name: homeassistant-video-call-updater.ps1
# Version 1.3 - Updated 2025-10-11
# Author: Michael Groat
# GitHub: https://github.com/mmgroat/home_automations/blob/main/homeassistant-video-call-updater.ps1
# Description: This script updates the color and brightness of smart lights based on video call status.
# It is designed to work with Home Assistant and smart lights that support color changes via the Home Assistant REST
# API. This script monitors for video call applications (like Zoom) running on the local machine. When a video call 
# application is detected, it changes the color and brightness of specified smart lights to an alert color (e.g.,
# bright pink) to indicate that the user is in a video call. When the video call application is no longer detected, it
# restores the lights to their previous state. The script uses the Home Assistant REST API to control the smart lights.
# Configuration settings are read from a JSON file. The script runs in an infinite loop, checking the state of the 
# specified processes every 30 seconds. Note: Ensure that the settings.json file contains the correct Home Assistant 
# base path, token, process names, entity IDs, and color/brightness settings.

# Usually, it is best to run this script as a scheduled task at logon with highest privileges. I find it's best to use
# lightbulbs that are located outside of the office to let others know you are in a call. Make sure the lights you use 
# are not used for other purposes, as this script will change their color and brightness when a call is detected. If
# the light is already on a different color when a call starts, it will change to the alert color and brightness, and 
# then revert back to the previous color and brightness when the call ends. If the light is in the alert color when
# the script starts, or not in a call, it will set the ligth to a default color and brightness.

# Load settings from JSON file
$settings_file = 'C:\Users\MikeGroat\OneDrive - Forge Global, Inc\Desktop\Personel\Mike_Groat\bin\settings.json'
$SettingsObject = Get-Content $settings_file | ConvertFrom-Json

# set up the URLs for the REST API calls
$onurl = $SettingsObject.ha_basepath + $SettingsObject.light_on_api
$offurl = $SettingsObject.ha_basepath + $SettingsObject.light_off_api
$geturlbase = $SettingsObject.ha_basepath + 'states/'

# set up the headers for the REST API calls
$headers = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
$headers.Add('Authorization', 'Bearer ' + $SettingsObject.ha_token)
$headers.Add('Content-Type', 'application/json')

# Global variables
$global:is_toggled = $false
$global:entity_last_states = New-Object 'System.Collections.Generic.Dictionary[[String],[Object]]'

function Get-Entity-State {
	param([string]$entity)

	# Get the current state of the entity
	$geturl = $geturlbase + $entity
	Start-Sleep -Milliseconds $SettingsObject.rest_API_delay_ms
	$output = Invoke-RestMethod $geturl -Method 'GET' -Headers $headers
	Write-Entity-State -entity $entity -output $output
	return $output
}

function Write-Entity-State {
	param([string]$entity, [Object]$output)

	if ($null -eq $output -or $null -eq $output.attributes) {
		[Console]::WriteLine("Could not retrieve state for $($entity)")
	} elseif ($null -eq $output.attributes.rgb_color -or $null -eq $output.attributes.brightness) {
		[Console]::WriteLine("Got state for $($entity): Light bulb is off")
	} else {
		[Console]::WriteLine("Got state for $($entity): color: $($output.attributes.rgb_color) and brightness: " + 
			"$($output.attributes.brightness)")
	}
	[Console]::Out.Flush()
}

function Set-Light-Color {
	param([string]$entity, [string]$color)	

	"Setting color of $entity to $color"
	$body = "{ `"entity_id`": `"$entity`", `"rgb_color`": $color }"
	Start-Sleep -Milliseconds $SettingsObject.rest_API_delay_ms
	Invoke-RestMethod $onurl -Method 'POST' -Headers $headers -Body $body
	Remove-Variable body
}

function Set-Light-Brightness {
	param([string]$entity, [string]$brightness)

	"Setting brightness of $entity to $brightness"
	$body = "{ `"entity_id`": `"$entity`", `"brightness`": $brightness }"
	Start-Sleep -Milliseconds $SettingsObject.rest_API_delay_ms
	Invoke-RestMethod $onurl -Method 'POST' -Headers $headers -Body $body
	Remove-Variable body
}

function Turn-Off-Light {
	param([string]$entity)

	"Turning off $entity"
	$body = "{ `"entity_id`": `"$entity`" }"
	Start-Sleep -Milliseconds $SettingsObject.rest_API_delay_ms
	Invoke-RestMethod $offurl -Method 'POST' -Headers $headers -Body $body
	Remove-Variable body
}

function Is-Alert-Color {
	param([Object]$output)	

	return ($null -ne $output -and $null -ne $output.attributes.rgb_color -and 
		# Check if the color is within the error range of the alert color. Sometimes it is off by a few values.
		($output.attributes.rgb_color[0] -le $SettingsObject.alert_color.r + 
		$SettingsObject.alert_color_error_range) -and 
		($output.attributes.rgb_color[0] -ge $SettingsObject.alert_color.r - 
		$SettingsObject.alert_color_error_range) -and 
		($output.attributes.rgb_color[1] -le $SettingsObject.alert_color.g + 
		$SettingsObject.alert_color_error_range) -and 
		($output.attributes.rgb_color[1] -ge $SettingsObject.alert_color.g - 
		$SettingsObject.alert_color_error_range) -and 
		($output.attributes.rgb_color[2] -le $SettingsObject.alert_color.b + 
		$SettingsObject.alert_color_error_range) -and 
		($output.attributes.rgb_color[2] -ge $SettingsObject.alert_color.b - 
		$SettingsObject.alert_color_error_range))
}

function Set-Entity-Last-State {
	param([string]$entity)

	$output = Get-Entity-State($entity)
	if (Is-Alert-Color($output)) {
		# The light is in the alert color. Set the last state to the default color.
		"Setting last state for $entity to default color and brightness"
		$global:entity_last_states[$entity]['color'] = $default_color_string
		$global:entity_last_states[$entity]['brightness'] = $SettingsObject.default_brightness
	} elseif ($null -eq $output -or $null -eq $output.attributes.rgb_color -or 
		$null -eq $output.attributes.brightness) {
		# The light bulb is turned off. Store the last sate as off (null values).
		"Setting last state for $entity to off (null values)"
		$global:entity_last_states[$entity]['color'] = $null
		$global:entity_last_states[$entity]['brightness'] = $null
	} else {
		# The light bulb is not in the alert color, nor turned off. Store color and brightness in last_entity_states.
		"Setting last state for $entity to color: $($output.attributes.rgb_color) and brightness: " + 
		"$($output.attributes.brightness)"
		$global:entity_last_states[$entity]['color'] = "[$($output.attributes.rgb_color[0]), "
		$global:entity_last_states[$entity]['color'] += "$($output.attributes.rgb_color[1]), "
		$global:entity_last_states[$entity]['color'] += "$($output.attributes.rgb_color[2])]"
		$global:entity_last_states[$entity]['brightness'] = $output.attributes.brightness.ToString()
	}
	Remove-Variable output
}

function Toggle-On {
	param([string]$entity)

	# Record the last state of the entity before changing it. Since this is called only once before changing the light
	# to the alert color (on a second method invocation while in the same call, is_toggled will be true), it will 
	# always store the last state before changing it to the alert color.
	if (! $global:is_toggled) {
		Set-Entity-Last-State $entity
	}
	# Always set the light to the alert color and alert brightness while in a call. Sometimes, someone changes the light
	# color while in a call. 
	Set-Light-Color $entity $alert_color_string
	Set-Light-Brightness $entity $SettingsObject.alert_brightness
}

function Toggle-Off {
	param([string]$entity)

	"Toggling off $entity (checking if light is in alert color, and restoring last state if so)"
	# Only change the light back if it is in the alert color. If someone changed the color while not in a call, leave it
	# alone. Always check if the light is in the alert color, because it could be in that state when the app started, or 
	# someone changed it to that color while not in a call, or we just got out of a call.
	if (Is-Alert-Color(Get-Entity-State($entity))) {
		# Restore the last state of the entity. If the last state was off (null), turn the light off. If the last state 
		# was the alert color, set it to the default color (but this is stored in the the entity_last_states variable, 
		# so always use the last state unless turning off the light).
		$last_color = $entity_last_states[$entity]['color']
		$last_brightness = $entity_last_states[$entity]['brightness']
		if ($null -eq $last_color -or $null -eq $last_brightness) {
			Turn-Off-Light($entity)
		} else {
			Set-Light-Color $entity $last_color
			Set-Light-Brightness $entity $last_brightness
		}
		Remove-Variable last_color
		Remove-Variable last_brightness
	}
}

function Update-Entities {
	param([string[]]$entities, [bool]$state)

	if ($state -eq $True) {
		foreach ($entity in $entities) {
			Toggle-On($entity)
		}
		$global:is_toggled = $true
	} else {
		foreach ($entity in $entities) {
			Toggle-Off($entity)
		}
		$global:is_toggled = $false
	}
}

function Check-Process {
	param([string]$processname, [string[]]$entities, [int]$offcallcount = 0)

	$process_var = Get-Process $processname -EA 0
	if ($process_var) {
		$processCount = (Get-NetUDPEndpoint -OwningProcess ($process_var).Id -EA 0 | Measure-Object).count
		if ($processCount -gt $offcallcount) {
			Update-Entities -entities $entities -state $True
		} else {    
			Update-Entities -entities $entities -state $False
		}
	} else {		
		Update-Entities -entities $entities -state $False
	}
	Remove-Variable process_var
}

# create the color strings for use in the REST calls
$alert_color_string = '[' + $SettingsObject.alert_color.r.ToString() + ', ' + $SettingsObject.alert_color.g.ToString()
$alert_color_string += ', ' + $SettingsObject.alert_color.b.ToString() + ']'
$default_color_string = '[' + $SettingsObject.default_color.r.ToString() + ', ' 
$default_color_string += $SettingsObject.default_color.g.ToString()
$default_color_string += ', ' + $SettingsObject.default_color.b.ToString() + ']'

$alert_color_string
$default_color_string

# intialize the entity_last_states values to default values
foreach ($process in $SettingsObject.processes) {
	foreach ($item in $SettingsObject.entities) {
		"Initializing last state for $item"
		$global:entity_last_states.Add($item, 
			@{ 'color' = $default_color_string; 'brightness' = $SettingsObject.default_brightness })
	}
}

# loop forever checking processes's states
while ($True) {
	'Checking processes at ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
	foreach ($process in $SettingsObject.processes) {
		Check-Process -processname $process.processname -entities $SettingsObject.entities -offcallcount `
			$process.nocallprocesscount
	}	
	Start-Sleep -Seconds 30
}
