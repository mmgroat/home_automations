########################################################################################################################
# Program Name: homeassistant-video-call-updater.ps1
# Version 1.4 - Updated 2025-10-17
# Author: Michael Groat
# GitHub: https://github.com/mmgroat/home_automations/blob/main/homeassistant-video-call-updater.ps1
#
# Description: This script updates the color and brightness of smart lights based on a video call status, or status of 
# other designated processes. It is designed to work with Home Assistant and smart lights that support color changes via 
# the Home Assistant REST API. This script monitors for video call applications (like Zoom) running on the local 
# machine. When a video call application is detected, it changes the color and brightness of specified smart lights to
# an alert color (e.g., bright pink) to indicate that the user is in a video call. When the video call application is no
# longer detected, it restores the lights to their previous state. Configuration settings are read from a JSON file. The
# script  runs in an infinite loop, checking the state of the specified processes every 30 seconds. Note: Ensure that 
# the settings.json file contains the correct Home Assistant base paths, token, process names, entity IDs, and 
# color/brightness settings.
#
# Usually, it is best to run this script as a scheduled task at logon. I find lightbulbs that are located outside of 
# the office let others know you are in a call. Make sure the lights you use are not used for other purposes, as this 
# script will change their color and brightness when a call is detected. If a light is on a non-alert color when a call 
# starts, it will change to the alert color and brightness, and then revert back to the previous color and brightness 
# when the call ends, or off if the light was off. If the light is in the alert color when the script starts, or when 
# not in a call, it will set the light to a default color and brightness specified in the JSON configuration file.
#
# Modifications:
# 2025-10-15 MMG Added temperature color mode ability
# 2025-11-20 MMG Added hs color mode ability
# 2025-12-12 MMG When turning off a light because the last state was off, change briefly the color and brightness first 
#                to the defaults. This prevents later, when turning on the light, turning it to the alert color and 
#                brightness, giving the false impression someone is on a call
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

# Get the current state of an entity (light bulb)	
function Get-Entity-State {
	param([string]$entity)

	$output = Invoke-Rest-API -url $($geturlbase + $entity) -method 'GET'
	Write-Entity-State -entity $entity -output $output
	return $output
}

# Log the state of the entity to the console
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
		} elseif ($output.attributes.color_mode -eq 'hs') {
			[Console]::WriteLine("Got state for $($entity): color mode: $($output.attributes.color_mode), " +
				"hs_color: $($output.attributes.hs_color), brightness: $($output.attributes.brightness)")
		} elseif ($output.attributes.color_mode -eq 'color_temp') {
			[Console]::WriteLine("Got state for $($entity): color mode: $($output.attributes.color_mode), " +
				"color_temp_kelvin: $($output.attributes.color_temp_kelvin), brightness: " +
				"$($output.attributes.brightness)")
		} else {
			[Console]::WriteLine("Error: Unknown color mode $($output.attributes.color_mode) for $entity")
		}
	}
	[Console]::Out.Flush()
}

# Set the light color by HS values
function Set-Light-Color-By-HS {
	param([string]$entity, [int[]]$color)	

	[Console]::WriteLine("Setting color by HS of $entity to $color")
	if ($color.Length -ne 2) {
		[Console]::WriteLine("Error: Color array must have exactly 2 elements (H, S). Given: $color")
	} else {
		$color_string = "[$($color[0]), $($color[1])]"
		$body = "{ `"entity_id`": `"$entity`", `"hs_color`": $color_string }"
		Invoke-Rest-API -url $onurl -method 'POST' -body $body
		Remove-Variable body
		Remove-Variable color_string
	}
}

# Set the light color by RGB values
function Set-Light-Color-By-RGB {
	param([string]$entity, [int[]]$color)	

	[Console]::WriteLine("Setting color by RGB of $entity to $color")
	if ($color.Length -ne 3) {
		[Console]::WriteLine("Error: Color array must have exactly 3 elements (R, G, B). Given: $color")
	} else {
		$color_string = "[$($color[0]), $($color[1]), $($color[2])]"
		$body = "{ `"entity_id`": `"$entity`", `"rgb_color`": $color_string }"
		Invoke-Rest-API -url $onurl -method 'POST' -body $body
		Remove-Variable body
		Remove-Variable color_string
	}
}

# Set the light color by temperature in Kelvin
function Set-Light-Color-By-Temp {
	param([string]$entity, [int]$temperature)

	[Console]::WriteLine("Setting color by temperature of $entity to $temperature")
	if ($null -eq $temperature) {
		[Console]::WriteLine("Error: Color temperature must be specified. Given: $temperature")
	} else {
		$body = "{ `"entity_id`": `"$entity`", `"color_temp_kelvin`": $($temperature.ToString()) }"
		Invoke-Rest-API -url $onurl -method 'POST' -body $body
		Remove-Variable body
	}
}

# Set the light brightness
function Set-Light-Brightness {
	param([string]$entity, [int]$brightness)

	[Console]::WriteLine("Setting brightness of $entity to $brightness")
	if ($null -eq $brightness) {
		[Console]::WriteLine("Error: Brightness must be specified. Given: $brightness")
	} else {
		$body = "{ `"entity_id`": `"$entity`", `"brightness`": $brightness }"
		Invoke-Rest-API -url $onurl -method 'POST' -body $body
		Remove-Variable body
	}
}

# Turn off the light
function Turn-Off-Light {
	param([string]$entity)

	[Console]::WriteLine("Turning off $entity")
	Invoke-Rest-API -url $offurl -method 'POST' -body "{ `"entity_id`": `"$entity`" }"
}

function Invoke-Rest-API {
	param([string]$url, [string]$method, [string]$body = $null)

	Start-Sleep -Milliseconds $SettingsObject.rest_API_delay_ms
	# Initialize the $output variable. This is needed in case of an exception from Invoke-RestMethod, otherwise the
	# previous value of $output would be used. (TODO: Look into Remove-Variable instead? But it is returned at the end
	# of the function)
	$output = $null 
	try {
		if ($method -eq 'POST') {
			$output = Invoke-RestMethod $url -Method $method -Headers $headers -Body $body
		} else {
			$output = Invoke-RestMethod $url -Method $method -Headers $headers
		}
	} catch [System.Net.WebException] {
		[Console]::WriteLine("Error: An exception occurred connecting to Home Assistant for $($entity): " +
			"$($_.Exception.Message)")
	}
	return $output
}	

# Check if the light is in the alert color
function Is-Alert-Color {
	param([Object]$output)	

	return ($null -ne $output -and 
		$null -ne $output.state -and 
		$output.state -eq 'on' -and 
		$null -ne $output.attributes -and 
		((Is-Alert-Color-RGB -output $output.attributes) -or 
			(Is-Alert-Color-HS -output $output.attributes)))
}

# Check if the light is in the alert color by RGB
function Is-Alert-Color-RGB {
	param([Object]$output)
	
	return ($output.color_mode -eq 'rgb' -and 
			$null -ne $output.rgb_color -and 
			# Check if the color is within the error range of the alert color. Sometimes it is off by a few values.
			((($output.rgb_color[0] -le $SettingsObject.alert_color[0] + 
					$SettingsObject.alert_color_error_range) -and 
				($output.rgb_color[0] -ge $SettingsObject.alert_color[0] - 
					$SettingsObject.alert_color_error_range) -and 
				($output.rgb_color[1] -le $SettingsObject.alert_color[1] + 
					$SettingsObject.alert_color_error_range) -and 
				($output.rgb_color[1] -ge $SettingsObject.alert_color[1] - 
					$SettingsObject.alert_color_error_range) -and 
				($output.rgb_color[2] -le $SettingsObject.alert_color[2] + 
					$SettingsObject.alert_color_error_range) -and 
				($output.rgb_color[2] -ge $SettingsObject.alert_color[2] - 
					$SettingsObject.alert_color_error_range)) -or 
				# Sometimes the color sticks to a different, but similar, alert color outside the error range,
	     		# so check for a mismatch color as well.
		    	(($output.rgb_color[0] -eq $SettingsObject.alert_color_mismatch[0]) -and 
					($output.rgb_color[1] -eq $SettingsObject.alert_color_mismatch[1]) -and 
					($output.rgb_color[2] -eq $SettingsObject.alert_color_mismatch[2]))))
}

# Check if the light is in the alert color by hs
function Is-Alert-Color-HS {
	param([Object]$output)

	return ($output.color_mode -eq 'hs' -and $null -ne $output.hs_color -and
			(($output.hs_color[0] -le $SettingsObject.alert_color_hs[0] + 
					$SettingsObject.alert_color_error_range) -and 
				($output.hs_color[0] -ge $SettingsObject.alert_color_hs[0] - 
					$SettingsObject.alert_color_hs_error_range) -and 
				($output.hs_color[1] -le $SettingsObject.alert_color_hs[1] + 
					$SettingsObject.alert_color_hs_error_range) -and 
				($output.hs_color[1] -ge $SettingsObject.alert_color_hs[1] - 
					$SettingsObject.alert_color_hs_error_range)))
}

# Set the last known state of the entity to the current state.
# This usually happens before changing the light to the alert color.
# If the light is already in the alert color, set the last state to the default color.
function Set-Entity-Last-State {
	param([string]$entity)

	[Console]::WriteLine("Setting the last state variable for $entity")	
	if (Is-Alert-Color($global:entity_last_states[$entity] = Get-Entity-State($entity))) {
		Set-Entity-Last-State-To-Default -entity $entity
	}
}

# Set the entity to the default color and brightness
function Set-Entity-Last-State-To-Default {
	param([string]$entity)

	[Console]::WriteLine("Setting the last state variable for $entity to the default temperature and brightness")
	if ($null -eq $global:entity_last_states[$entity] -or $null -eq $global:entity_last_states[$entity].attributes) {
		$global:entity_last_states[$entity] = {
			attributes: {
				color_temp_kelvin: $SettingsObject.default_temperature,
				brightness: $SettingsObject.default_brightness,
				color_mode: 'color_temp'
			}
		}
	} else {
		$global:entity_last_states[$entity].attributes.color_temp_kelvin = $SettingsObject.default_temperature
		$global:entity_last_states[$entity].attributes.brightness = $SettingsObject.default_brightness
		$global:entity_last_states[$entity].attributes.color_mode = 'color_temp'
	}
}

# Set the entity to the alert color and brightness
function Toggle-On {
	param([string]$entity)

	# Record the last state of the entity before changing it. Since this is called only once before changing the light
	# to the alert color (on a second method invocation while in the same call, is_toggled will be true), it will 
	# always store the last state before changing it to the alert color.
	if (! $global:is_toggled) {
		Set-Entity-Last-State -entity $entity
	}
	# _Always_ set the light to the alert color and alert brightness while in a call. Sometimes, someone changes the
	# light color while in a call. It is just as expensive as retreiving the state, so just always set it. 
	Set-Light-Color-By-RGB -entity $entity -color $SettingsObject.alert_color
	Set-Light-Brightness -entity $entity -brightness $SettingsObject.alert_brightness
}

# Restore the last known state of the entity if it is in the alert color
function Toggle-Off {
	param([string]$entity)

	[Console]::WriteLine("Toggling off $entity (checking if light is in alert color, and restoring last state if so)")
	# Only change the light back if it is in the alert color. If someone changed the color while not in a call, leave
	# it alone. _Always_ check if the light is in the alert color, because it could be in that state when the app
	# started, or someone changed it to that color while not in a call, or, most likely, we just got out of a call.
	if (Is-Alert-Color(Get-Entity-State -entity $entity )) {
		# Restore the last state of the entity. If the last state state was off, turn the light off. If the last state
		# was the alert color, set it to the default color (but this is stored in the the entity_last_states variable,
		# anyways, so always use the last state variable, unless turning off the light).
		if ($null -eq $entity_last_states[$entity] -or $null -eq $entity_last_states[$entity].state -or
			$entity_last_states[$entity].state -eq 'off' -or $null -eq $entity_last_states[$entity].attributes) {
			# Once in a while, the last state will be off, and then the light is changed to the alert color because of
			# a call. In these cases, when the light is later turned back on after it has turned off (even if this
			# script has stopped running), the light turns on to the alert color. To resolve this, we briefly turn the
			# light to the default color and brightness before turning it off. This way, if the light is turned on
			# later, even after the script has stopped running, the light does not enter the alert color, giving the
			# wrong impression someone is on a call.
			Set-Light-Brightness -entity $entity -brightness $SettingsObject.default_brightness
			Set-Light-Color-By-Temp -entity $entity -temp $SettingsObject.default_temperature
			Turn-Off-Light -entity $entity
		} else {
			Set-Light-Brightness -entity $entity -brightness $entity_last_states[$entity].attributes.brightness
			if ($entity_last_states[$entity].attributes.color_mode -eq 'rgb') {
				Set-Light-Color-By-RGB -entity $entity -color $entity_last_states[$entity].attributes.rgb_color
			} elseif ($entity_last_states[$entity].attributes.color_mode -eq 'hs') {
				Set-Light-Color-By-HS -entity $entity -color $entity_last_states[$entity].attributes.hs_color
			} elseif ($entity_last_states[$entity].attributes.color_mode -eq 'color_temp') {
				Set-Light-Color-By-Temp -entity $entity -temp $entity_last_states[$entity].attributes.color_temp_kelvin
			} else {
				[Console]::WriteLine("Error: Unknown color mode " + 
					"$($entity_last_states[$entity].attributes.color_mode) for $entity")
			}
		}
	}
}

# Update the entities to the specified state (on or off)
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

# Check if the specified process is running and update the entities accordingly
function Check-Process {
	param([string]$processname, [string[]]$entities, [int]$offcallcount = 0)

	$process_var = Get-Process $processname -EA 0
	if ($process_var) {
		$processCount = (Get-NetUDPEndpoint -OwningProcess ($process_var).Id -EA 0 | Measure-Object).count
		"Process $processname is running with $processCount Net UDP endpoints."
		if ($processCount -gt $offcallcount) {
			Update-Entities -entities $entities -state $True
		} else {    
			Update-Entities -entities $entities -state $False
		}
		Remove-Variable processCount
	} else {		
		Update-Entities -entities $entities -state $False
	}
	Remove-Variable process_var
}

# intialize the entity_last_states values to default values
foreach ($entity in $SettingsObject.entities) {
	"Initializing last state for $entity"
	$global:entity_last_states.Add($entity, $(Get-Entity-State -entity $entity))
	Set-Entity-Last-State-To-Default -entity $entity
}

# loop forever checking processes's states
while ($True) {
	'Checking processes at ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
	foreach ($process in $SettingsObject.processes) {
		Check-Process -processname $process.processname -entities $SettingsObject.entities -offcallcount `
			$process.nocallprocesscount
	}	
	Get-Process -Name powershell | Sort-Object -Descending WS | ` 
		Select-Object -First 10 Name, @{Name = 'WorkingSetMB'; Expression = { $_.WorkingSet / 1MB } }
	Start-Sleep -Seconds 30
}
