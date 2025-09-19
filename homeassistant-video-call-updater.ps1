$SettingsObject = Get-Content "C:\Users\MikeGroat\OneDrive - Forge Global, Inc\Desktop\Personel\Mike_Groat\bin\settings.json" | ConvertFrom-Json
$onurl = $SettingsObject.openhabbasepath + 'services/light/turn_on'
$offurl = $SettingsObject.openhabbasepath + 'services/light/turn_off'
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization", "Bearer " + $SettingsObject.openhabtoken)
$headers.Add("Content-Type", "application/json")
$is_toggled = $false
$global:entity_last_states = New-Object "System.Collections.Generic.Dictionary[[String],[Object]]"
#in case the program starts with the lights turned to red, then default to this color and brightness - we don't usually have them red anyways
$default_color = "[255, 235, 218]"
$default_brightness = "51" 

function Update-OpenHAB {
	param([string]$item, [bool]$state)

	if ($state -eq $True) {
		toggle_on("light.mike_s_light")
		toggle_on("light.living_room_light_four")
		$global:is_toggled = $true
	}
	else
	{
		toggle_off("light.mike_s_light")
		toggle_off("light.living_room_light_four")
		$global:is_toggled = $false
	}
}

function toggle_on {
	param([string]$entity)
	
	set_entity_last_states($entity)
	turn_red($entity)
	set_full_brightness($entity)
}

function set_entity_last_states {
	param([string]$entity)

	if (! $global:is_toggled) {
		$output = get_entity_state($entity)
		if (is_red($output)){
			# app was started with light bulbs red - set last state to default color
			"app was started with light bulbs red - set to default color"
			$global:entity_last_states[$entity]["color"] = $default_color
			$gloabl:entity_last_states[$entity]["brightness"] = $default_brightness
		} elseif ($output -eq $null -or $output.attributes.rgb_color -eq $null -or $output.attributes.brightness -eq $null) {
			# light bulb is turned off
			"Light bulbs were turned off"
			$global:entity_last_states[$entity]["color"] = $null
			$global:entity_last_states[$entity]["brightness"] = $null
		} else {
			#light bulb is not red - store previous state
			$red = $output.attributes.rgb_color[0]
			$green = $output.attributes.rgb_color[1]
			$blue = $output.attributes.rgb_color[2]
			$global:entity_last_states[$entity]["color"] = "[$red, $green, $blue]"
			$global:entity_last_states[$entity]["brightness"] = $output.attributes.brightness
		}
	}
}

function turn_red {
	param([string]$entity)
	
	$body = "{ `"entity_id`": `"$entity`", `"rgb_color`": [255, 0, 0] }"
	Start-Sleep -Milliseconds 1000
	Invoke-RestMethod $onurl -Method 'POST' -Headers $headers -Body $body
}

function set_full_brightness {
	param([string]$entity)
	
	$body = "{ `"entity_id`": `"$entity`", `"brightness_pct`": 100 }"
	Invoke-RestMethod $onurl -Method 'POST' -Headers $headers -Body $body
}

function toggle_off {
	param([string]$entity)
	
	$output = get_entity_state($entity)
	if (is_red($output)) {
		"in toggle off is red"
		$entity_last_color = $global:entity_last_states[$entity]["color"]
		$entity_last_brightness = $global:entity_last_states[$entity]["brightness"]
		if($entity_last_color -eq $null -or $entity_last_brightness -eq $null){
			"Turning off entity"
			$body = "{ `"entity_id`": `"$entity`" }"
			Start-Sleep -Milliseconds 1000
			Invoke-RestMethod $offurl -Method 'POST' -Headers $headers -Body $body
		}
		else
		{
			$body1 = "{ `"entity_id`": `"$entity`", `"rgb_color`": $entity_last_color }"
			$body1
			Start-Sleep -Milliseconds 1000
			Invoke-RestMethod $onurl -Method 'POST' -Headers $headers -Body $body1
			$body2 = "{ `"entity_id`": `"$entity`", `"brightness`": $entity_last_brightness }"
			$body2
			Start-Sleep -Milliseconds 1000
			Invoke-RestMethod $onurl -Method 'POST' -Headers $headers -Body $body2
		}
	}
}

function is_red {
	param([Object]$output)
	
	return (($output -ne $null) -and ($output.attributes.rgb_color -ne $null) -and 
		($output.attributes.rgb_color[0] -ge 224 -and $output.attributes.rgb_color[1] -eq 0 -and $output.attributes.rgb_color[2] -eq 0 ))
}

function get_entity_state {
	param([string]$entity)
	
	$geturl = $SettingsObject.openhabbasepath + "states/$entity"
	$output = Invoke-RestMethod $geturl -Method 'GET' -Headers $headers
	
	return $output
}

function Check-Process {
	param([string]$processname, [string]$openhabitem, [int]$offcallcount = 0)
	
	$process = Get-Process $processname -EA 0

	if($process) {
		$processCount = (Get-NetUDPEndpoint -OwningProcess ($process).Id -EA 0|Measure-Object).count
		
		if ($processCount -gt 5) {
			Update-OpenHAB -item $openhabitem -state $True
		}
		else {    
			Update-OpenHAB -item $openhabitem -state $False
		}
	}
	else {		
		Update-OpenHAB -item $openhabitem -state $False
	}
	Remove-Variable process
}

$global:mike_s_light = @{ "color" = $default_color; "brightness" = $default_brightness }
$global:living_room_light_four = @{ "color" = $default_color; "brightness" = $default_brightness }
$global:entity_last_states.Add("light.mike_s_light",$mike_s_light)
$global:entity_last_states.Add("light.living_room_light_four",$living_room_light_four)

While($True) {

	Foreach ($process in $SettingsObject.processes) {
		"Waking up - checking lights"
		Check-Process -processname $process.processname -openhabitem $process.openhabitem -offcallcount $process.nocallprocesscount
	}
	
	Start-Sleep -Seconds 30
}