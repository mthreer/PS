$timer=[system.diagnostics.stopwatch]::StartNew()

###################################################
# get current script path
###################################################
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$SavedLocation = Get-Location
$logPath = "$scriptPath\LicenseCheckLogs"
if (-not(Test-Path $logPath)) {
	New-Item -Type Directory $logPath
}

###################################################
# logging to screen and file
###################################################
$datetime = Get-Date -format "yyyy-MM-dd_HH.mm"
function Write-Logging($logtext, $fg, $bg, $Line, $ScreenOnly, $FileOnly) {
	$paramColor = @{}
	if ($bg) { $paramColor.Add("BackgroundColor", "$bg" ) }
	if ($fg) { $paramColor.Add("ForegroundColor", "$fg" ) }
	if ($line) { if (-not($FileOnly) ) { Write-Host "$logtext" -NoNewLine @paramColor } }
	
	# log to screen
	if (-not($line)) { if (-not($FileOnly) ) { Write-Host "$logtext" @paramColor } }
	# log to file
	if (-not($ScreenOnly) ) {
		Add-Content "$logPath\LicenseCheck_$datetime.txt" "$logtext"
	}
	if ($bg) { $paramColor.Remove("BackgroundColor") }
	if ($fg) { $paramColor.Remove("ForegroundColor") }
}

$QueryUser=Read-Host "Input UserPrincipalName to check or 'all' for everyone"

if ($QueryUser -ieq "all") {
	"Querying every user, this will take some time"
	$Input=Get-MsolUser -All -ErrorAction SilentlyContinue | select objectid,userprincipalname,licenses
}
if (-not($QueryUser -ieq "all")) {
	$Input=Get-MsolUser -UserPrincipalName $QueryUser -ErrorAction SilentlyContinue | select objectid,userprincipalname,licenses
	if (-not($Input)) {
		Write-Host "Could not find that user: $QueryUser" -foregroundColor Red
		exit
	}
}

Write-Logging -logtext ("{0};{1};{2};{3};{4};{5}" -f 'User.ObjectID','UserPrincipalName','SkuID','AssignedDirectly','AssignedFromGroup','GroupName(s)') -FileOnly "1"
Write-Logging -logtext ("{0,-60}{1,-42}{2,-20}{3,-20}{4}" -f 'UserPrincipalName','SkuID','AssignedDirectly','AssignedFromGroup','GroupName(s)') -fg Cyan -ScreenOnly "1"

$count = 1; $PercentComplete = 0; $script:countinner = 1
foreach ($User in $Input) {
	#Progress message
	$ActivityMessage = "Retrieving license status for user $($User.UserPrincipalName). Please wait..."
	$StatusMessage = ("Processing {0} of {1}: {2}" -f $count, @($Input).count, $($User.UserPrincipalName))
	$PercentComplete = ($count / @($Input).count * 100)
	Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
	$count++

	$Licenses=$User.Licenses
	$GroupName=$null
	#foreach ($License in $Licenses.GetEnumerator() | sort-object -Property $_) {
	foreach ($License in $Licenses) {
		$AssignedDirectly=$False
		$AssignedFromGroup=$False
		$GroupName=$null
		
		# GroupsAssigningLicense contains a collection of IDs of objects assigning the license
		# This could be a group object or a user object (contrary to what the name suggests)
		# If the collection is empty, this means the license is assigned directly - this is the case for users who have never been licensed via groups in the past
		if ($License.GroupsAssigningLicense.Count -eq 0) {
			$AssignedDirectly="Yes"
			$AssignedFromGroup="Never"
			
			#Write-Logging -logtext ("{0};{1};{2};{3};{4};{5}" -f $($User.ObjectID),$($User.UserPrincipalName),$($License.AccountSkuID),$AssignedDirectly,$AssignedFromGroup,$GroupName) -FileOnly "1"
			#Write-Logging -logtext ("{0,-40}{1,-42}{2,-20}{3,-20}{4}" -f $($User.UserPrincipalName),$($License.AccountSkuID),$AssignedDirectly,$AssignedFromGroup,$GroupName) -fg Cyan -ScreenOnly "1"
		}

		# If collection is not empty, this means the user has been or is being, assigned license directly OR inheriting licenses from a group
		if ($License.GroupsAssigningLicense.Count -ge 1) {
		
			# GroupsAssigningLicense contains a collection of IDs of objects assigning the license
			# This could be a group object or a user object (contrary to what the name suggests)
			foreach ($LicenseGroupID in $License.GroupsAssigningLicense) {
				# If the current ObjectID in the property GroupsAssigningLicense does NOT equal the ObjectID of the user object, this means the user is inheriting license from a group
				if (-not($LicenseGroupID -ieq $User.ObjectID)) {
					# Try to retrieve the groups displayname
					Try {
						# If we haven't already identified a group for this License assignment, do this now
						if (-not($GroupName)) {
							$GroupName=Get-MsolGroup -ObjectId $LicenseGroupID -ErrorAction Stop | select -ExpandProperty DisplayName
						}
						# If we have already identified a group for this License assignment, identify the next one
						if ($GroupName) {
							$Other=Get-MsolGroup -ObjectID $LicenseGroupID -ErrorAction Stop | select -ExpandProperty DisplayName
							# Prevent identifying duplicate groups
							# If the $Other group doesnt already exist in the GroupName collection, append it - otherwise keep already identified group in $GroupName as is.
							if (-not($GroupName -eq $Other)) {
								$GroupName+=" / $Other"
							}
						}
					}
					# If we could not retrieve the groups DisplayName, set GroupName to the ObjectID that failed to be identified.
					Catch {
						# If empty, set to ID of unidentified group
						if (-not($GroupName)) {
							$GroupName="$LicenseGroupID"
						}
						# Append unidentified groups to already identified groups
						if ($GroupName) {
							$GroupName+=" / $LicenseGroupID"
						}
					}
					$AssignedFromGroup=$True
				}
				# If the current ObjectID in the property GroupsAssigningLicense does equal the ObjectID of the user object, this means the user has been assigned license directly.
				if ($LicenseGroupID -ieq $User.ObjectID) {
					$AssignedDirectly=$True
					# $GroupName=$null
				}
			}
		}
		Write-Logging -logtext ("{0};{1};{2};{3};{4};{5}" -f $($User.ObjectID),$($User.UserPrincipalName),$($License.AccountSkuID),$AssignedDirectly,$AssignedFromGroup,$GroupName) -FileOnly "1"
		Write-Logging -logtext ("{0,-60}{1,-42}{2,-20}{3,-20}{4}" -f $($User.UserPrincipalName),$($License.AccountSkuID),$AssignedDirectly,$AssignedFromGroup,$GroupName) -fg Cyan -ScreenOnly "1"
	}
	
	# Pause progress bar for 2 seconds before ending loop (for visuality)
	if ($PercentComplete -eq 100) {
		Start-sleep 2
	}
}
$Timer.Stop()
Write-Logging -logtext ""
Write-Logging -logtext "Script took $($Timer.Elapsed.Hours)h $($Timer.Elapsed.Minutes)m $($Timer.Elapsed.Seconds)s $($Timer.Elapsed.MilliSeconds)ms to run." -fg Yellow