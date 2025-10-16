########################################################################################################################
# Program Name: homeassistant-video-call-updater.ps1
# Version 1.3 - Updated 2025-10-11
# Author: Michael Groat
# GitHub: https://github.com/mmgroat/home_automations/blob/main/homeassistant-video-call-updater.ps1
#
# Description: This script updates the color and brightness of smart lights based on video call status.
# It is designed to work with Home Assistant and smart lights that support color changes via the Home Assistant REST
# API. This script monitors for video call applications (like Zoom) running on the local machine. When a video call 
# application is detected, it changes the color and brightness of specified smart lights to an alert color (e.g.,
# bright pink) to indicate that the user is in a video call. When the video call application is no longer detected, it
# restores the lights to their previous state. The script uses the Home Assistant REST API to control the smart lights.
# Configuration settings are read from a JSON file. The script runs in an infinite loop, checking the state of the 
# specified processes every 30 seconds. Note: Ensure that the settings.json file contains the correct Home Assistant 
# base path, token, process names, entity IDs, and color/brightness settings.
#
# Usually, it is best to run this script as a scheduled task at logon with highest privileges. I find it's best to use
# lightbulbs that are located outside of the office to let others know you are in a call. Make sure the lights you use 
# are not used for other purposes, as this script will change their color and brightness when a call is detected. If
# the light is already on a different color when a call starts, it will change to the alert color and brightness, and 
# then revert back to the previous color and brightness when the call ends. If the light is in the alert color when
# the script starts, or not in a call, it will set the ligth to a default color and brightness.
#
# Modifications:
# 2025-10-15 MMG Added color temperature color mode ability
########################################################################################################################

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
	try {
		$output = Invoke-RestMethod $geturl -Method 'GET' -Headers $headers
	} catch [System.Net.WebException] {
		[Console]::WriteLine("An exception occurred connecting to Home Assistant for entity $($entity): " + 
			"$($_.Exception.Message)")
	}
	Remove-Variable geturl
	Write-Entity-State -entity $entity -output $output
	return $output
}

function Write-Entity-State {
	param([string]$entity, [Object]$output)

	if ($null -eq $output -or $null -eq $output.attributes -or $null -eq $output.state) {
		[Console]::WriteLine("Could not retrieve state for $($entity)")
	} elseif ($output.state -eq 'off') {
		[Console]::WriteLine("Got state for $($entity): Light bulb is off")
	} else {
		if ($output.attributes.color_mode -eq 'rgb') {
			[Console]::WriteLine("Got state for $($entity): color mode: $($output.attributes.color_mode), " +
				"rgb_color: $($output.attributes.rgb_color), brightness: $($output.attributes.brightness)")
		} else {
			[Console]::WriteLine("Got state for $($entity): color mode: $($output.attributes.color_mode), " +
				"color_temp_kelvin: $($output.attributes.color_temp_kelvin), brightness: " +
				"$($output.attributes.brightness)")
		}
	}
	[Console]::Out.Flush()
}

function Set-Light-Color-By-RGB {
	param([string]$entity, [int[]]$color)	

	"Setting color by RGB of $entity to $color"
	if ($color.Length -ne 3) {
		[Console]::WriteLine("Error: Color array must have exactly 3 elements (R, G, B). Given: $color")	
	} else {
		$color_string = "[$($color[0]), $($color[1]), $($color[2])]"
		$body = "{ `"entity_id`": `"$entity`", `"rgb_color`": $color_string }"
		Start-Sleep -Milliseconds $SettingsObject.rest_API_delay_ms
		Invoke-RestMethod $onurl -Method 'POST' -Headers $headers -Body $body
		Remove-Variable body
		Remove-Variable color_string
	}
}

function Set-Light-Color-By-Temp {
	param([string]$entity, [int]$temperature)	

	"Setting color by temperature of $entity to $temperature"
	$body = "{ `"entity_id`": `"$entity`", `"color_temp_kelvin`": $($temperature.ToString()) }"
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

	return ($null -ne $output -and $null -ne $output.state -and $output.state -eq 'on' -and 
		$null -ne $output.attributes -and $output.attributes.color_mode -eq 'rgb' -and
		$null -ne $output.attributes.rgb_color -and 
		# Check if the color is within the error range of the alert color. Sometimes it is off by a few values.
		($output.attributes.rgb_color[0] -le $SettingsObject.alert_color[0] + 
		$SettingsObject.alert_color_error_range) -and 
		($output.attributes.rgb_color[0] -ge $SettingsObject.alert_color[0] - 
		$SettingsObject.alert_color_error_range) -and 
		($output.attributes.rgb_color[1] -le $SettingsObject.alert_color[1] + 
		$SettingsObject.alert_color_error_range) -and 
		($output.attributes.rgb_color[1] -ge $SettingsObject.alert_color[1] - 
		$SettingsObject.alert_color_error_range) -and 
		($output.attributes.rgb_color[2] -le $SettingsObject.alert_color[2] + 
		$SettingsObject.alert_color_error_range) -and 
		($output.attributes.rgb_color[2] -ge $SettingsObject.alert_color[2] - 
		$SettingsObject.alert_color_error_range))
}

function Set-Entity-Last-State {
	param([string]$entity)

	$global:entity_last_states[$entity] = Get-Entity-State($entity)
	if (Is-Alert-Color($global:entity_last_states[$entity])) {
		# The light is in the alert color. Set the last state to the default color.
		Set-Entity-To-Default $entity
	}
}

function Set-Entity-To-Default {
	param([string]$entity)

	"Setting last state for $entity to default temperture and brightness"
	$global:entity_last_states[$entity].attributes.color_temp_kelvin = $SettingsObject.default_temperature
	$global:entity_last_states[$entity].attributes.brightness = $SettingsObject.default_brightness
	$global:entity_last_states[$entity].attributes.color_mode = 'color_temp'
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
	Set-Light-Color-By-RGB $entity $SettingsObject.alert_color
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
		if ($null -eq $entity_last_states[$entity] -or $null -eq $entity_last_states[$entity].state -or
			$entity_last_states[$entity].state -eq 'off' ) {
			Turn-Off-Light $entity
		} else {
			Set-Light-Brightness $entity $entity_last_states[$entity].attributes.brightness
			if ($entity_last_states[$entity].attributes.color_mode -eq 'rgb') {
				Set-Light-Color-By-RGB $entity $entity_last_states[$entity].attributes.rgb_color
			} else {
				Set-Light-Color-By-Temp $entity $entity_last_states[$entity].attributes.color_temp_kelvin
			}
		}
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

# intialize the entity_last_states values to default values
foreach ($entity in $SettingsObject.entities) {
	"Initializing last state for $entity"
	$output = Get-Entity-State($entity)
	$global:entity_last_states.Add($entity, $output)
	Set-Entity-To-Default $entity
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
