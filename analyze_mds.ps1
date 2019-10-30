# Script for analyzing and outputting MDS configuration
# Joshua Woleben
# Written 9/18/2019

# Load SSH

Import-Module -Name "networkshare\Powershell\Modules\Posh-SSH.psm1"
Import-Module -Name "networkshare\Powershell\Modules\Posh-SSH.psd1"

$terminal_command = @'
terminal length 0
'@

$running_config_command = @'
show running-config
'@

$flogi_command = @'
show flogi database
'@

$password = ""
# Create Excel file

$excel_file = "$env:USERPROFILE\Documents\CiscoMDSLog_$(get-date -f MMddyyyyHHmmss).xlsx"
# Open Excel

# Create new Excel object
$excel_object = New-Object -comobject Excel.Application
$excel_object.visible = $True 

$mds_hosts = @("MDS1","MDS2")

$fc_array = @()
$peer_array = @()
$path_array = @()
$zone_array = @()
$member_array = @()

# Create new Excel workbook
$excel_workbook = $excel_object.Workbooks.Add()
$worksheet_item = 1
$excel_worksheet = $excel_workbook.Worksheets.Add()

foreach ($MDSHostname in $mds_hosts) {
    # Select the first worksheet in the new workbook
    
    $excel_worksheet.Name = $MDSHostname

    # Write-Host ($zone_array[$i].name + "," + $member + "," + $fc_label + "," + $peer_wwn)
    # Create headers
    $excel_worksheet.Cells.Item(1,1) = "Zone Name"
    $excel_worksheet.Cells.Item(1,2) = "Member Name"
    $excel_worksheet.Cells.Item(1,3) = "FibreChannel Port Label"
    $excel_worksheet.Cells.Item(1,4) = "Peer WWN"

    # Format headers
    $d = $excel_worksheet.UsedRange

    # Set headers to backrgound pale yellow color, bold font, blue font color
    $d.Interior.ColorIndex = 19
    $d.Font.ColorIndex = 11
    $d.Font.Bold = $True

    # Set first data row
	$row_counter = 2

	$session = ""
	$flogi_output = ""
	$running_config_output = ""
	$running_config = ""
	$running_config_stream = ""
	$flogi_output = ""
	$flogi_stream = ""
	$creds = ""


	$fc_array = @()
	$peer_array = @()
	$path_array = @()
	$zone_array = @()
	$member_array = @()

	$username = "admin"
	$password = Read-Host -Prompt "Enter the password for the $username user" -AsSecureString


	$creds = New-Object -TypeName System.Management.Automation.PSCredential ($username,$password)
	Write-Host "Connecting to $MDSHostname..."

	$session = New-SSHSession -ComputerName $MDSHostname -Credential $creds
	$flogi_stream = New-SSHShellStream -SessionId $session.SessionId
	Invoke-SSHStreamShellCommand -ShellStream $flogi_stream -Command $terminal_command -PrompPattern "#"
	sleep 5
	$flogi_output = Invoke-SSHStreamShellCommand -ShellStream $flogi_stream -Command $flogi_command
	sleep 5
	Remove-SSHSession -SessionId $session.SessionId

	sleep 5

	$session = New-SSHSession -ComputerName $MDSHostname -Credential $creds
	 sleep 3 
	$running_config_stream = New-SSHShellStream -SessionId $session.SessionId

		$running_config_stream.Write("terminal length 0`n")
		sleep 5
		$running_config_stream.Write("show running-config`n")
		sleep 15
		$running_config = $running_config_stream.Read()



	Remove-SSHSession -SessionId $session.SessionId

	sleep 5
	Write-Host "Parsing FLOGI output..."
	Write-Host $flogi_output
	Write-Host $running_config
	($flogi_output | Select-String -AllMatches -CaseSensitive -Pattern "(?smi)(.+)\s+?(0x.+?)\s+?([a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2})\s+?([a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2})").Matches | ForEach-Object {
			$current_fc_label = $_.Groups[1].Value
			$current_fcid = $_.Groups[2].Value
			$current_fc_wwn = $_.Groups[3].Value
			$current_peer_wwn = $_.Groups[4].Value
			$fc_object = New-Object psobject -Property @{
				label = ($current_fc_label -replace "\s+\d+?\s+","").Trim()
				fcid = $current_fcid.Trim()
				wwn = $current_fc_wwn.Trim()
				peer_wwn = $current_peer_wwn.Trim()
			}
			Write-Host $fc_object
			$fc_array += $fc_object
	}

	# Get peer port data
	Write-Host "Analyzing WWN Aliases..."
	$peer_data = ($running_config | Select-String -AllMatches -Pattern "(?smi)device-alias name .+? pwwn [a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}").Matches

	For ($i = 1; $i -lt ($peer_data.Groups).Count; $i++) {
		$peer_data.Groups[$i].Value | Select-String -Pattern "device-alias name (.+?) pwwn ([a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2})" | ForEach-Object {
			$current_alias = $_.Matches.Groups[1].Value
			$current_wwn = $_.Matches.Groups[2].Value
			$peer_object = New-Object psobject -Property @{
				alias =  $current_alias.Trim()
				wwn = $current_wwn.Trim()
			}
			Write-Host $peer_object
			$peer_array += $peer_object
		}

	}


	# Get fiber paths
	Write-Host "Analyzing WWN and port associations..."
	$path_data =($running_config | Select-String -AllMatches -CaseSensitive -Pattern "(?smi)vsan \d+ wwn [a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2} fcid 0x[a-z0-9]{6}").Matches

	For($i = 1; $i -lt ($path_data.Groups).Count; $i++) {
		$path_data.Groups[$i].Value | Select-String -Pattern "vsan \d+ wwn ([a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}:[a-z0-9]{2}) fcid (0x[a-z0-9]{6})" | ForEach-Object {
			$peer_wwn = $_.Matches.Groups[1].Value
			$fcid = $_.Matches.Groups[2].Value

			$path_object = New-Object psobject -Property @{
				peer_wwn = $peer_wwn
				fcid = $fcid
			}
			Write-Host $path_object
			$path_array += $path_object

		}
		
	}

	# Get zone data
	Write-Host "Analyzing zones..."
	$zone_data = ($running_config| Out-String | Select-String -AllMatches -CaseSensitive -Pattern "(?smi)zone name .+? vsan.+?zone").Matches

	For ($i = 1; $i -lt ($zone_data.Groups).Count; $i++) {
		$zone_name = ($zone_data.Groups[$i].Value | Select-String -Pattern "(?smi)zone name (.+?) ").Matches.Groups[1].Value
		$member_data =($zone_data.Groups[$i].Value | Select-String -AllMatches -Pattern "(?smi)member device-alias (.+?) ").Matches
		Write-Host $member_data
		For ($j = 0; $j -lt ($member_data.Groups).Count; $j++) {
			$member_array =@()
			$member_data.Groups[$j].Value | Select-String -Pattern "member device-alias (.+?) " | ForEach-Object {
				$member = $_.Matches.Groups[1].Value
				$member_array += $member.Trim()
			}
			$zone_object = New-Object psobject -Property @{
				name = $zone_name.Trim()
				members = $member_array
			}
			Write-Host $zone_object
			$zone_array += $zone_object
		}
	}


	# Calculate port paths
	Write-Host "Analyzing port and device alias associations..."
	$fc_peer_object = ""
	$fc_peer_array = @()
	for ($i = 0; $i -lt $fc_array.Count; $i++) {
		$member_alias = ""
		$fc_peer_wwn = ($fc_array[$i] | Select -ExpandProperty wwn)
		for ($j = 0; $j -lt $peer_array.Count; $j++) {
			if ($fc_peer_wwn -match ($peer_array[$j] | Select -ExpandProperty wwn)) {
			   $member_alias = ($peer_array[$j] | Select -ExpandProperty alias)
			   $fc_peer_object = New-Object -TypeName PSObject -Property @{
					fc_label = ($fc_array[$i] | Select -ExpandProperty label)
					peer_label = $member_alias
					peer_wwn = $fc_peer_wwn

			   }
			   Write-Host $fc_peer_object
			   $fc_peer_array += $fc_peer_object

			}
		}
					 
	}

	Write-Host "Data below!"

	for ($i = 0; $i -lt $zone_array.Count; $i++ ) {
		foreach ($member in $zone_array[$i].members) {
		$fc_label = ""
		$peer_wwn = ""
			foreach ($fc_peer in $fc_peer_array) {
				if ($fc_peer.peer_label -match $member) {
					$fc_label = $fc_peer.fc_label
					$peer_wwn = $fc_peer.peer_wwn
				}
			}
		$excel_worksheet.Cells.Item($row_counter,1) = $zone_array[$i].name
		$excel_worksheet.Cells.Item($row_counter,2) = $member
		$excel_worksheet.Cells.Item($row_counter,3) = $fc_label
		$excel_worksheet.Cells.Item($row_counter,4) = $peer_wwn
		$row_counter++

		}
          
        

	}
    $e = $excel_worksheet.Range("A1:D$row_counter")
    $e.Borders.Item(12).Weight = 2
    $e.Borders.Item(12).LineStyle = 1
    $e.Borders.Item(12).ColorIndex = 1

    $e.Borders.Item(11).Weight = 2
    $e.Borders.Item(11).LineStyle = 1
    $e.Borders.Item(11).ColorIndex = 1

    # Set thicker border around outside
    $e.BorderAround(1,4,1)

    # Fit all columns
    $e.Columns("A:F").AutoFit()
    $worksheet_item++

    sleep 5
    $excel_worksheet = $excel_workbook.Worksheets.Add()
} # end of loop

    # Save Excel
    $excel_workbook.SaveAs($excel_file) | out-null

    # Quit Excel
    $excel_workbook.Close | out-null
    $excel_object.Quit() | out-null