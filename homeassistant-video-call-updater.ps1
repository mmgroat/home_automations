$SettingsObject = Get-Content "C:\Users\MikeGroat\OneDrive - Forge Global, Inc\Desktop\Personel\Mike_Groat\bin\settings.json" | ConvertFrom-Json
$onurl = $SettingsObject.openhabbasepath + 'services/light/turn_on'
$offurl = $SettingsObject.openhabbasepath + 'services/light/turn_off'
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization", "Bearer " + $SettingsObject.openhabtoken)
$headers.Add("Content-Type", "application/json")
$is_toggled = $false
$entity_last_states = New-Object "System.Collections.Generic.Dictionary[[String],[Object]]"
$default_color = "[255, 235, 218]"
$default_brightness = "51" # 20 percent of 255
$red_color = "[255, 0, 0]"
$full_brightness = "255"

function Update-Entities {
	param([string[]]$entities, [bool]$state)

	if ($state -eq $True) {
		Foreach ($entity in $entities) {
			toggle_on($entity)
		}
		$global:is_toggled = $true
	} else {
		Foreach($entity in $entities) {
			toggle_off($entity)
		}
		$global:is_toggled = $false
	}
}

function get_entity_state {
	param([string]$entity)
	
	$geturl = $SettingsObject.openhabbasepath + "states/$entity"
	$output = Invoke-RestMethod $geturl -Method 'GET' -Headers $headers	
	return $output
}

function set_light_color {
	param([string]$entity, [string]$color)
	
	$body = "{ `"entity_id`": `"$entity`", `"rgb_color`": $color }"
	Start-Sleep -Milliseconds 1000
	Invoke-RestMethod $onurl -Method 'POST' -Headers $headers -Body $body
}

function set_light_brightness {
	param([string]$entity, [string]$brightness)
	
	$body = "{ `"entity_id`": `"$entity`", `"brightness`": $brightness }"
	Start-Sleep -Milliseconds 1000
	Invoke-RestMethod $onurl -Method 'POST' -Headers $headers -Body $body
}

function turn_off_light {
	param([string]$entity)
	
	$body = "{ `"entity_id`": `"$entity`" }"
	Start-Sleep -Milliseconds 1000
	Invoke-RestMethod $offurl -Method 'POST' -Headers $headers -Body $body
}

function toggle_on {
	param([string]$entity)
	
	set_entity_last_states($entity)
	set_light_color $entity  $red_color 
	set_light_brightness $entity $full_brightness
}

function toggle_off {
	param([string]$entity)

	$output = get_entity_state($entity)
	if (is_red($output)) {
		$last_color = $global:entity_last_states[$entity]["color"]
		$last_brightness = $global:entity_last_states[$entity]["brightness"]
		if($last_color -eq $null -or $last_brightness -eq $null){
			turn_off_light($entity)
		} else {
			set_light_color $entity $last_color
			set_light_brightness $entity $last_brightness
		}
	}
}

function set_entity_last_states {
	param([string]$entity)

	if (! $global:is_toggled) {
		$output = get_entity_state($entity)
		if (is_red($output)){
			# the app was started with light bulb red - set last state to default color
			$global:entity_last_states[$entity]["color"] = $default_color
			$global:entity_last_states[$entity]["brightness"] = $default_brightness
		} elseif ($output -eq $null -or $output.attributes.rgb_color -eq $null -or $output.attributes.brightness -eq $null) {
			# when video was turned on, light bulb was turned off
			$global:entity_last_states[$entity]["color"] = $null
			$global:entity_last_states[$entity]["brightness"] = $null
		} else {
			#light bulb is not red nor turned off - store previous state
			$red = $output.attributes.rgb_color[0]
			$green = $output.attributes.rgb_color[1]
			$blue = $output.attributes.rgb_color[2]
			$global:entity_last_states[$entity]["color"] = "[$red, $green, $blue]"
			$global:entity_last_states[$entity]["brightness"] = $output.attributes.brightness.ToString()
		}
	}
}

function is_red {
	param([Object]$output)
	
	return (($output -ne $null) -and ($output.attributes.rgb_color -ne $null) -and 
		($output.attributes.rgb_color[0] -ge 224 -and $output.attributes.rgb_color[1] -eq 0 -and $output.attributes.rgb_color[2] -eq 0 ))
}

function Check-Process {
	param([string]$processname, [string[]]$entities, [int]$offcallcount = 0)
	
	$process = Get-Process $processname -EA 0

	if($process) {
		$processCount = (Get-NetUDPEndpoint -OwningProcess ($process).Id -EA 0|Measure-Object).count
		
		if ($processCount -gt 5) {
			Update-Entities -entities $entities -state $True
		}
		else {    
			Update-Entities -entities $entities -state $False
		}
	}
	else {		
		Update-Entities -entities $entities -state $False
	}
	Remove-Variable process
}

# intialize the entity_last_states values to default values
Foreach ($process in $SettingsObject.processes) {
	Foreach ($item in $SettingsObject.processes.entities) {
		$entity_last_states.Add($item, @{ "color" = $default_color; "brightness" = $default_brightness })
	}
}

# loop forever checking process's states
While($True) {

	Foreach ($process in $SettingsObject.processes) {
		Check-Process -processname $process.processname -entities $process.entities -offcallcount $process.nocallprocesscount
	}
	
	Start-Sleep -Seconds 30
}