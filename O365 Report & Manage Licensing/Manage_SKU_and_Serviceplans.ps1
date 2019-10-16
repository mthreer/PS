# Author: Niklas Jumlin (niklas.jumlin@atea.se)
# Version: 2019-01-31
# Copyright: Free to use, please leave this header intact 
#
# Requires connection to: MSOnline
#
# Install-Module MSOnline -MinimumVersion 1.1.183.17
# Version 1.1.183.17 allows for viewing which services are explicitly disabled via a SKU being licensed from a group
#

param(  
    [Parameter(
        Position=0, 
        Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [String]$UPN
)

# Check if module requirements are fulfilled
if ( -not((Get-Module -ListAvailable MSOnline).Version -ge "1.1.183.17") ) { "Module: MSOnline is not installed or the version is not 1.1.183.17 or greater" ; exit }

$timer=[system.diagnostics.stopwatch]::StartNew()

# Consider the below services as critical services that should be taken into consideration before
# removing a SKU-pack. The script will compare these services towards other group enabled services in any other SKU-packs. 
# If the SKU-pack to be removed is the only source of licensing for these services, then recommend that the service should not be removed.

$DoNotIgnore=@(
	"EXCHANGE_S_STANDARD"
	"EXCHANGE_S_ENTERPRISE"
	"EXCHANGE_S_DESKLESS"
	"SHAREPOINTSTANDARD_EDU"
	"SHAREPOINTSTANDARD"
	"SHAREPOINTENTERPRISE"
	"SHAREPOINTENTERPRISE_EDU"
	"SHAREPOINTDESKLESS"
	"TEAMS1"
)

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$logPath = "$scriptPath\LicenseData"
if (-not(Test-Path $logPath)) {
	"$logpath does not exist, creating it"
	New-Item -Type Directory $logPath | Out-Null
}

function GetDT {
	$script:dt = Get-Date -format "yyyy-MM-dd HH:mm:ss"
}

###################################################
# logging to screen and file
###################################################
$datetime = Get-Date -format "yyyy-MM-dd_HH.mm"
function Write-Logging($logtext, $fg, $bg, [switch]$NoNewLine, [switch]$ScreenOnly, [switch]$FileOnly) {
	$paramColor = @{}
	if ($bg) { $paramColor.Add("BackgroundColor", "$bg" ) }
	if ($fg) { $paramColor.Add("ForegroundColor", "$fg" ) }
	if ($NoNewLine) { if (-not($FileOnly) ) { Write-Host "$logtext" -NoNewLine @paramColor } }

	# log to screen
	if (-not($NoNewLine)) { if (-not($FileOnly) ) { Write-Host "$logtext" @paramColor } }
	# log to file
	if (-not($ScreenOnly) ) {
		Add-Content "$logPath\$UPN`_$datetime`_LicenseData.txt" "$logtext"
	}
	if ($bg) { $paramColor.Remove("BackgroundColor") }
	if ($fg) { $paramColor.Remove("ForegroundColor") }
}

Try {
	$User=Get-MsolUser -UserPrincipalName $UPN -ErrorAction Stop | select-object DisplayName, UserPrincipalName, ObjectID, Licenses, Title, Department, LicenseAssignmentDetails
	#  | select -ExpandProperty LicenseAssignmentDetails | select -ExpandProperty Assignments | select -ExpandProperty ReferencedObjectId | Select -ExpandProperty Guid
	#LicenseAssignmentDetails.Assignments.ReferencedObjectId.Guid
}
Catch {
	$ErrorMessage=$_.Exception.Message
	Write-Logging -logtext "$ErrorMessage" -fg Red
	exit
}

if ($User) {
	Write-Logging -logtext "Checking licenses towards $($User.DisplayName) `($($User.UserPrincipalName)`) with objectID $($User.ObjectID)" -fg Green
	Write-Logging -logtext "Title: $($User.Title) Department: $($User.Department)" -fg Green
	Write-Logging -logtext ""
	# Create a list of all GUIDs (groups or user) that assign licenses to the user
	[array]$GroupsAssigningLicense=@()
	foreach ($GroupID in $User.Licenses.GroupsAssigningLicense) {
		[array]$GroupsAssigningLicense += $GroupID.GUID
	}
	# Remove duplicates
	[array]$GroupsAssigningLicense = $GroupsAssigningLicense | sort-object -Unique

	# collect all referenced object IDs
	foreach ($ReferencedObjectId in $User.LicenseAssignmentDetails.Assignments.ReferencedObjectId.Guid) {
		if ($ReferencedObjectId -notin $GroupsAssigningLicense) {
			[array]$PendingAssignmentIds += $ReferencedObjectId
			[array]$GroupsAssigningLicense += $ReferencedObjectId
		}
	}

	Write-Logging -logtext "Objects assigning licenses:" -fg Gray
	Write-Logging -logtext ""
	# Collect all disabled plans per group
	$SkuPerGroup=@{}

	# kept for debugging:
	#$DisabledServices=@{}


	# $DisabledServicesPerSkuPerGroup stores data like this:
	# @{"Group" = @{ "SkuPack" = @(Disabled Services Array) } }
	# 
	# Which resembles data like a tree structure
	# + Group (Key)
	#		+ SkuPack (Key)
	#			+	Disabled Services (array)
	$DisabledServicesPerSkuPerGroup=@{}
	foreach ($GroupID in $GroupsAssigningLicense) {
			# if its the users objectid
			if ($GroupID -eq $User.ObjectID) {
				$GroupName="$($User.UserPrincipalName)"
			}
			if (-not($GroupID -eq $User.ObjectID)) {
				$GrpSkuPacks=@()
				$GroupLicenseData=Get-MsolGroup -ObjectId $GroupID | Select-Object DisplayName, AssignedLicenses
				$GroupName=$GroupLicenseData.DisplayName

				$DisabledServicesPerSkuPerGroup.Add("$GroupName",$null)

				# kept for debugging:
				#[array]$DisabledServicePlans=@()
				foreach ($LicenseData in $GroupLicenseData.AssignedLicenses) {
					# kept for debugging:
					#[array]$rewrittenPlan=@()
					foreach ($Plan in $LicenseData) {
						$Sku="$($LicenseData.AccountSkuId.AccountName):$($LicenseData.AccountSkuId.SkuPartNumber)"
						# Create a collection of all SkuPacks being applied from groups
						$GrpSkuPacks+=$Sku
						$DisabledServicesPerSkuPerGroup[$GroupName] += @{"$Sku"=@($Plan.DisabledServicePlans)}
						
						# kept for debugging:
						#foreach ($DisabledService in $Plan.DisabledServicePlans) { [Array]$rewrittenPlan += "$Sku - $DisabledService" }
						#if (($Plan.DisabledServicePlans).Count -eq 0) { [Array]$rewrittenPlan += "$Sku - None" }
					}
					# kept for debugging:
					#$DisabledServicePlans+=$rewrittenPlan
				}
				# kept for debugging:
				#$DisabledServices.Add("$GroupName",$DisabledServicePlans)
			}
			Write-Logging -logtext "$GroupName `($GroupID`)" -fg DarkGray
			$SkuPerGroup.Add("$GroupName",$GrpSkuPacks)
	}
	# filter out duplicate skupacks from collection
	$GrpSkuPacks=@()
	foreach ($val in $SkuPerGroup.Values) {
			$GrpSkuPacks+=$val
	}
	$GrpSkuPacks = $GrpSkuPacks | sort-object -Unique

	# Collect all available services per SKU
	# This will be used to compare against per group disabled services in order to figure out which services are left enabled
	$AllSkuServices=@{}
	foreach ($SkuPack in $GrpSkuPacks) {
		$AllSkuServices.Add("$SkuPack",$null)
		$SkuServices=(Get-MsolAccountSku | where {$_.AccountSkuId -eq $SkuPack}).ServiceStatus.ServicePlan.ServiceName
		foreach ($Service in $SkuServices) {
			[Array]$AllSkuServices[$SkuPack] += $Service
		}
	}
	Write-Logging -logtext ""
	Write-Logging -logtext "Per Group Disabled/Enabled Services:" -fg Gray
	Write-Logging -logtext ""
	$DisabledServicesCollection=@()
	foreach ($Group in $DisabledServicesPerSkuPerGroup.Keys) {
		Write-Logging -logtext "$Group" -fg Gray
		foreach ($SkuName in $AllSkuServices.Keys) {
			if ($SkuName -in $DisabledServicesPerSkuPerGroup[$Group].Keys) {
				foreach ($Service in $AllSkuServices[$SkuName]) {
					if ($Service -in $DisabledServicesPerSkuPerGroup[$Group][$SkuName]) {
						[Array]$DisabledServicesCollection += "$($SkuName):$($Service)"
						Write-Logging -logtext ("{0,-70}{1}" -f "$($SkuName):$($Service)",'Disabled') -fg DarkGray
					}
					else {
						Write-Logging -logtext ("{0,-70}{1}" -f "$($SkuName):$($Service)",'Enabled') -fg DarkGray
					}
				}
				Write-Logging -logtext ""
			}
		}
	}

	Write-Logging -logtext "SkuPacks from groups:" -fg Gray
	Write-Logging -logtext ""
	foreach ($SkuPack in $GrpSkuPacks) {
		Write-Logging -logtext "$SkuPack" -fg DarkGray
	}
	Write-Logging -logtext ""
	Write-Logging -logtext "SkuPacks directly licensed on user" -fg Gray
	Write-Logging -logtext ""
	$DirectSkuPacks=@()
	# if the number of SkuPacks with ObjectIDs belonging to the User itself is 0, then the user has no direct licenses obviously
	if ($User.Licenses) {
		if (-not($User.Licenses.GroupsAssigningLicense)) {
			foreach ($UserData in $User.licenses) {
				Write-Logging -logtext "$($UserData.AccountSkuID)" -fg DarkGray
				[array]$DirectSkuPacks += $UserData.AccountSkuID
			}
		}
		if ($User.Licenses.GroupsAssigningLicense) {
			if ( ($User.licenses | Where-Object {$_.GroupsAssigningLicense -eq $User.ObjectID}).Count -eq 0) {
				Write-Logging -logtext "None" -fg DarkGray
			}
			else {
				foreach ($UserData in $User.licenses | Where-Object {$_.GroupsAssigningLicense -eq $User.ObjectID}) {
					Write-Logging -logtext "$($UserData.AccountSkuID)" -fg DarkGray
					[array]$DirectSkuPacks += $UserData.AccountSkuID
				}
			}
		}
	}
	Write-Logging -logtext ""
	if ($User.Licenses) {
		Write-Logging -logtext ("{0,-45}{1,-5}{2,-50}{3,-23}{4,-25}{5,-18}{6}" -f 'LicensePack (SkuID)','', 'Service','Status','GroupAssignedService','Source','Groups') -ScreenOnly
		Write-Logging -logtext ("{0,-45}{1,-5}{2,-50}{3,-23}{4,-25}{5,-18}{6}" -f '-------------------','', '-------','------','--------------------','------','------') -ScreenOnly
		Write-Logging -logtext ("{0,-45};{1,-50};{2,-23};{3,-25};{4,-18};{5}" -f 'LicensePack (SkuID)', 'Service','Status','GroupAssignedService','Source','Groups') -FileOnly
		$count=0
		$Debug=@()
		$ServiceValidationPerSku=@{}
		$DirectEnabledServicesPerSku=@{}
		$ExtraEnabledServicesPerSku=@{}
		foreach ($UserLicense in $User.Licenses) {
			$count+=1
			$SkuName=$UserLicense.AccountSkuID
			$Services=$UserLicense.ServiceStatus
			if ($count -ge 2) {
				# new line (separate output per SkuID)
				Write-Logging -logtext ""
			}

			$SourceCollection=@()
			$DirectEnabledServices=@()
			$ExtraEnabledServices=@()

			$countservices=1
			foreach ($Service in $Services) {
				$ServiceName=$Service.ServicePlan.ServiceName
				$ServiceType=$Service.ServicePlan.ServiceType
				$ServiceDisplayName="$($ServiceName) `($($ServiceType)`)"
				$ServiceStatus=$Service.ProvisioningStatus

				# Reset variables before each iteration
				$DisablingGroups=@()
				$EnablingGroups=@()
				$Source=""
				$Color=""
				$BGColor=""
				$Color2=""

				# $DisabledServicesPerSkuPerGroup stores data like this:
				# @{"Group" = @{ "SkuPack" = @(Disabled Services Array) } }
				# 
				# Which resembles data like a tree structure
				# + Group
				#		+ SkuPack
				#			+	Disabled Services
				
				# loop all groups and their disabled services per skupack
				foreach ($Group in $DisabledServicesPerSkuPerGroup.Keys) {
					# if the current SkuPack being validated equals that of the group disabled skupack
					if ($UserLicense.AccountSkuID -in $DisabledServicesPerSkuPerGroup[$Group].Keys) {
						# if the service exists in the array of disabled services for this skupack and group
						if ($ServiceName -in $DisabledServicesPerSkuPerGroup[$Group][$UserLicense.AccountSkuID]) {
							# append this specific group that is disabling this service to an array
							[Array]$DisablingGroups+=$Group
							# kept for debugging:
							#Write-Host "$($UserLicense.AccountSkuID):$($ServiceName) was found Disabled in $Group" -ForegroundColor Red
						}
						else {
							# append this specific group enabling this service to an array
							[Array]$EnablingGroups+=$Group
							# kept for debugging:
							#Write-Host "$($UserLicense.AccountSkuID):$($ServiceName) was found Enabled in $Group" -ForegroundColor Green
						}
					}
					else {
						# [Array]$EnablingGroups+=$Group
					}
				}
				foreach ($PendingGroupId in $PendingGroupIds) {
					# if $ServiceName -in 
				}
				# Fix array (array to string, separated with comma)
				$DisablingGroupsString=$DisablingGroups -join ", "
				$EnablingGroupsString=$EnablingGroups -join ", "
				
				# Evaluation of source

				# If there is at least 1 groups enabling this service, the source is Group
				if ( ($EnablingGroups).Count -gt 0) {
					$GroupAssignedService="Enabled"
					$Source="Group"
				}

				# If there is no groups enabling this service and at least one group disabling it - this service, if enabled, is considered an Extra (Direct) licensed service.
				if ( (($EnablingGroups).Count -eq 0) -and (($DisablingGroups).Count -gt 0) ) {
					$GroupAssignedService="Disabled"
					if ( ($ServiceStatus -eq "Success") -or ($ServiceStatus -eq "PendingInput") -or ($ServiceStatus -eq "PendingActivation") -or ($ServiceStatus -eq "PendingProvisioning") ) {
						$Source="Extra (Direct)"
					}
				}

				# If there are no groups enabling or disabling this service and if the SkuPack is directly licensed, this specific service is considered directly licensed if its enabled.
				if ( (($EnablingGroups).Count -eq 0) -and (($DisablingGroups).Count -eq 0) ) {
					$GroupAssignedService="False"
					if ($UserLicense.AccountSkuID -in $DirectSkuPacks) {
						if ( ($ServiceStatus -eq "Success") -or ($ServiceStatus -eq "PendingInput") -or ($ServiceStatus -eq "PendingActivation") -or ($ServiceStatus -eq "PendingProvisioning") ) {
							$Source="Direct"
						}
					}
				}
				# If this service is a GroupAssignedService (validated above) and if the SkuPack is directly licensed, this service is licensed from both Group and PROBABLY Direct if its enabled.
				if ( ($UserLicense.AccountSkuID -in $DirectSkuPacks) -and ($GroupAssignedService -eq "Enabled") -and ( ($ServiceStatus -eq "Success") -or ($ServiceStatus -eq "PendingInput") -or ($ServiceStatus -eq "PendingActivation") -or ($ServiceStatus -eq "PendingProvisioning") ) ) {
					$Source="Direct+Group"
				}

				# kept for debugging:
				#"Disabling: $DisablingGroups"
				#"Enabling: $EnablingGroups"

				# only output SkuName for the first iteration
				if (-not($countservices -eq "1")) {
					$SkuName=$null
				}

				# Print output in different colors for every other row (odd/even)
				$countservices | ForEach-Object {
					if ($_ % 2 -eq 1 ) {
						$Color2="White"
					}
					else {
						$Color2="DarkGray"
					}
				}
				Write-Logging -logtext ("{0,-45}{1,-5}{2,-50}{3,-23}{4,-25}" -f $SkuName,':', $ServiceDisplayName, $ServiceStatus, $GroupAssignedService) -ScreenOnly -NoNewLine -fg $Color2

				if ($Source -eq "Direct") {
					$Color="Yellow"
					$BGColor="Black"
				}
				if ($Source -eq "Group") {
					$Color="Gray"
					$BGColor=""
				}
				if ($Source -eq "Direct+Group") {
					$Color="DarkRed"
					$BGColor="Black"
				}
				if ($Source -eq "Extra (Direct)") {
					$Color="Red"
					$BGColor="Black"
				}
				if (-not($Source)) {
					$Color="Green"
					$BGColor=""
				}
				if ($Source -eq "Extra (Direct)") {
					Write-Logging -logtext ("{0,-18}" -f $Source) -fg $Color -bg $BGColor -ScreenOnly

					# file only logging
					Write-Logging -logtext ("{0,-45};{1,-50};{2,-23};{3,-25};{4}" -f $($UserLicense.AccountSkuID), $ServiceDisplayName, $ServiceStatus, $GroupAssignedService, $Source) -FileOnly
				}
				if ( ($Source -eq "Direct+Group") -or ($Source -eq "Group") ) {
					Write-Logging -logtext ("{0,-18}" -f $Source) -fg $Color -bg $BGColor -ScreenOnly -NoNewLine
					Write-Logging -logtext ("{0}" -f $EnablingGroupsString) -fg $Color2 -ScreenOnly

					# file only logging
					Write-Logging -logtext ("{0,-45};{1,-50};{2,-23};{3,-25};{4,-18};{5}" -f $($UserLicense.AccountSkuID), $ServiceDisplayName, $ServiceStatus, $GroupAssignedService, $Source, $EnablingGroupsString) -FileOnly
				}
				if ( ($Source -eq "Direct") -or (-not($Source)) ) {
					Write-Logging -logtext ("{0}" -f $Source) -fg $Color -bg $BGColor -ScreenOnly

					# file only logging
					Write-Logging -logtext ("{0,-45};{1,-50};{2,-23};{3,-25};{4}" -f $($UserLicense.AccountSkuID), $ServiceDisplayName, $ServiceStatus, $GroupAssignedService, $Source) -FileOnly
				}

				# Keep count of services (purely for output)
				$countservices+=1

				if ($Source) {
					[array]$SourceCollection += @{"$ServiceName"="$Source"}
				}

				if ( ($ServiceStatus -eq "Success") -or ($ServiceStatus -eq "PendingInput") -or ($ServiceStatus -eq "PendingActivation") -or ($ServiceStatus -eq "PendingProvisioning") ) {
					if ($Source -eq "Direct") {
						[array]$DirectEnabledServices += $ServiceName
					}
					if ($Source -eq "Extra (Direct)") {
						[array]$ExtraEnabledServices += $ServiceName
					}
					if (($Source -eq "Group") -or ($Source -eq "Direct+Group")) {
						[array]$GroupEnabledService += $ServiceName
					}
				}

				if ($GroupAssignedService -eq "Disabled") {
					[Array]$Debug+="$($UserLicense.AccountSkuID):$ServiceName was found disabled in: $DisablingGroupsString"
				}
			}
			$ServiceValidationPerSku.Add("$($UserLicense.AccountSkuID)",$SourceCollection)
			$DirectEnabledServicesPerSku.Add("$($UserLicense.AccountSkuID)",$DirectEnabledServices)
			$ExtraEnabledServicesPerSku.Add("$($UserLicense.AccountSkuID)",$ExtraEnabledServices)
		}

		Write-Logging -logtext ""
		Write-Logging -logtext "DEBUG:" -fg Yellow
		foreach ($row in $Debug) {
			Write-Logging -logtext "$Row" -fg Yellow
		}

		# remove licenses
		#
		# ServiceValidationPerSku is a HashTable that contains a HashTable:
		# @{"SkuID"= @{"Service"="Extra (Direct)"} }
		#
		# Visualized into the following:
		# "SkuID"
		# 	"Service"="Extra (Direct)"
		#
		foreach ($skuid in $ServiceValidationPerSku.Keys) {
			[array]$extraServices=@()
			[array]$directServices=@()
			# Key contains the service name, the value contains the source, e.g Extra (Direct) or Direct etc.
			foreach ($Svc in $ServiceValidationPerSku[$skuid].Keys) {
				if ($ServiceValidationPerSku[$skuid].$Svc -eq "Extra (Direct)") {
					[array]$extraServices += $Svc
				}
				if ($ServiceValidationPerSku[$skuid].$Svc -eq "Direct") {
					[array]$directServices += $Svc
				}
			}

			# Validate Extra (Direct) services being enabled (redundant or not)
			if ("Extra (Direct)" -in $ServiceValidationPerSku[$skuid].Values) {
				Write-Logging -logtext ""
				Write-Logging -logtext "Validating Extra (Direct) licensed services in SKU: $skuid" -fg Magenta
				Write-Logging -logtext ""
				$RemoveSku=$skuid
				foreach ($ExtraSvc in $ExtraEnabledServicesPerSku[$skuid]) {
					Write-Logging -logtext ("{0,-30}" -f $ExtraSvc) -fg Cyan -NoNewline -ScreenOnly
					if ($ExtraSvc -in $DoNotIgnore) {
						Write-Logging -logtext ("{0,-18}" -f '(Critical)') -fg Red -NoNewLine -ScreenOnly
						if ($ExtraSvc -notin $GroupEnabledService) {
							Write-Logging -logtext ("{0,-30}{1,-18}{2,-20}{3,20}{4}" -f $ExtraSvc,'(Critical)','Is NOT redundant', $skuid, ' can NOT be removed') -FileOnly
							Write-Logging -logtext ("{0,-20}" -f 'Is NOT redundant') -fg Red -NoNewLine -ScreenOnly
							Write-logging -logtext ("{0,20}{1}" -f $skuid,' can NOT be removed') -fg Red -ScreenOnly
							$RemoveSku=$False
							[array]$CriticalServices += $ExtraSvc
							# kept for debugging:
							#Break;
						}
						if ($ExtraSvc -in $GroupEnabledService) {
							Write-Logging -logtext ("{0,-30}{1,-18}{2,-20}{3,20}{4}" -f $ExtraSvc,'(Critical)','Is redundant', $skuid, ' can be removed') -FileOnly
							Write-Logging -logtext ("{0,-20}" -f 'Is redundant') -fg Green -NoNewLine -ScreenOnly
							Write-logging -logtext ("{0,20}{1}" -f $skuid,' can be removed') -fg Green -ScreenOnly
							[array]$CriticalRedundantServices += $ExtraSvc
						}
					}
					else {
						Write-Logging -logtext ("{0,-18}" -f '(Not critical)') -fg Green -NoNewLine -ScreenOnly
						if ($DirectSvc -notin $GroupEnabledService) {
							Write-Logging -logtext ("{0,-30}{1,-18}{2,-20}{3,20}{4}" -f $ExtraSvc,'(Not critical)','Is NOT redundant', $skuid, ' can be removed') -FileOnly
							Write-Logging -logtext ("{0,-20}" -f 'Is NOT redundant') -fg DarkRed -NoNewLine -ScreenOnly
							Write-logging -logtext ("{0,20}{1}" -f $skuid,' can be removed') -fg Yellow -ScreenOnly
							[array]$SafeServices += $ExtraSvc
						}
						if ($DirectSvc -in $GroupEnabledService) {
							Write-Logging -logtext ("{0,-30}{1,-18}{2,-20}{3,20}{4}" -f $ExtraSvc,'(Not critical)','Is redundant', $skuid, ' can be removed') -FileOnly
							Write-Logging -logtext ("{0,-20}" -f 'Is redundant') -fg Green -NoNewLine -ScreenOnly
							Write-logging -logtext ("{0,20}{1}" -f $skuid,' can be removed') -fg Green -ScreenOnly
							[array]$SafeRedundantServices += $ExtraSvc
						}
					}
				}
				""
				"Result: RemoveSku: $RemoveSku"

				if (-not($RemoveSku -eq $False)) {
					$Timer.Stop()
					Write-Logging -logtext ""
					Write-Logging -logtext "Continue removing $skuid`? (Default is No)"
					$Readhost = Read-Host "( y / n )" 
					$Timer.Start()
					Switch ($ReadHost) { 
						Y { $RemoveExtra=$True } 
						J { $RemoveExtra=$True } 
						N { $RemoveExtra=$False } 
						Default { $RemoveExtra=$False } 
					}
					if ($RemoveExtra -eq $True) {
						Write-Logging -logtext ""
						Write-Logging -logtext "Removing $skuid from $UPN" -fg Cyan -bg Black
						Try {
							Set-MsolUserLicense -ObjectId $($User.ObjectID) -RemoveLicenses $skuid -ErrorAction Stop
						}
						Catch {
							$ErrorMessage=$_.Exception.Message
							Write-Logging -logtext "$ErrorMessage" -fg Red
						}
					}

					if ($RemoveExtra -eq $False) {
						Write-Logging -logtext ""
						Write-Logging -logtext "Skipping removal of $skuid on $UPN" -fg DarkGray
						Add-Content "$logPath\Skipped.csv" "$UPN;$skuid;Direct;$($extraServices -join ', ');$($CriticalServices -join ', ');$($CriticalRedundantServices -join ', ');$($SafeServices -join ', ');$($SafeRedundantServices -join ', ')"					
					}
				}

				if ($RemoveSku -eq $False) {
					$Timer.Stop()
					Write-Logging -logtext ""
					Write-Logging -logtext "Would you like to override above and continue removing $skuid`? (Default is No)"
					$Readhost = Read-Host "( y / n )" 
					$Timer.Start()
					Switch ($ReadHost) { 
						Y { $OverrideRemoveExtra=$True } 
						J { $OverrideRemoveExtra=$True } 
						N { $OverrideRemoveExtra=$False } 
						Default { $OverrideRemoveExtra=$False } 
					}
					if ($OverrideRemoveExtra -eq $True) {
						Write-Logging -logtext ""
						Write-Logging -logtext "Removing $skuid from $UPN" -fg Cyan -bg Black
						Try {
							Set-MsolUserLicense -ObjectId $($User.ObjectID) -RemoveLicenses $skuid -ErrorAction Stop
						}
						Catch {
							$ErrorMessage=$_.Exception.Message
							Write-Logging -logtext "$ErrorMessage" -fg Red
						}
					}

					if ($OverrideRemoveExtra -eq $False) {
						Write-Logging -logtext ""
						Write-Logging -logtext "Skipping removal of $skuid on $UPN (Removal of license would cause critical service loss)" -fg DarkGray
						Add-Content "$logPath\Skipped.csv" "$UPN;$skuid;Direct;$($extraServices -join ', ');$($CriticalServices -join ', ');$($CriticalRedundantServices -join ', ');$($SafeServices -join ', ');$($SafeRedundantServices -join ', ')"					
					}
				}
			}
			# end extra (direct) part

			# Validate directly licensed SKU (redundant or not)
			if ($GroupEnabledService.count -gt 0) {
				if ("Direct" -in $ServiceValidationPerSku[$skuid].Values) {
					Write-Logging -logtext ""
					Write-Logging -logtext "Validating Directly licensed SKU: $skuid" -fg Magenta
					Write-Logging -logtext ""
					$RemoveSku=$skuid
					foreach ($DirectSvc in $DirectEnabledServicesPerSku[$skuid]) {
						Write-Logging -logtext ("{0,-30}" -f $DirectSvc) -fg Cyan -NoNewline -ScreenOnly
						if ($DirectSvc -in $DoNotIgnore) {
							Write-Logging -logtext ("{0,-18}" -f '(Critical)') -fg Red -NoNewLine -ScreenOnly
							if ($DirectSvc -notin $GroupEnabledService) {
								Write-Logging -logtext ("{0,-30}{1,-18}{2,-20}{3,20}{4}" -f $DirectSvc,'(Critical)','Is NOT redundant', $skuid, ' can NOT be removed') -FileOnly
								Write-Logging -logtext ("{0,-20}" -f 'Is NOT redundant') -fg Red -NoNewLine -ScreenOnly
								Write-logging -logtext ("{0,20}{1}" -f $skuid,' can NOT be removed') -fg Red -ScreenOnly
								$RemoveSku=$False
								[array]$CriticalServices += $DirectSvc
								# kept for debugging:
								#Break;
							}
							if ($DirectSvc -in $GroupEnabledService) {
								Write-Logging -logtext ("{0,-30}{1,-18}{2,-20}{3,20}{4}" -f $DirectSvc,'(Critical)','Is redundant', $skuid, ' can be removed') -FileOnly
								Write-Logging -logtext ("{0,-20}" -f 'Is redundant') -fg Green -NoNewLine -ScreenOnly
								Write-logging -logtext ("{0,20}{1}" -f $skuid,' can be removed') -fg Green -ScreenOnly
								[array]$CriticalRedundantServices += $DirectSvc
							}
						}
						else {
							Write-Logging -logtext ("{0,-18}" -f '(Not critical)') -fg Green -NoNewLine -ScreenOnly
							if ($DirectSvc -notin $GroupEnabledService) {
								Write-Logging -logtext ("{0,-30}{1,-18}{2,-20}{3,20}{4}" -f $DirectSvc,'(Not critical)','Is NOT redundant', $skuid, ' can be removed') -FileOnly
								Write-Logging -logtext ("{0,-20}" -f 'Is NOT redundant') -fg DarkRed -NoNewLine -ScreenOnly
								Write-logging -logtext ("{0,20}{1}" -f $skuid,' can be removed') -fg Yellow -ScreenOnly
								[array]$SafeServices += $DirectSvc
							}
							if ($DirectSvc -in $GroupEnabledService) {
								Write-Logging -logtext ("{0,-30}{1,-18}{2,-20}{3,20}{4}" -f $DirectSvc,'(Not critical)','Is redundant', $skuid, ' can be removed') -FileOnly
								Write-Logging -logtext ("{0,-20}" -f 'Is redundant') -fg Green -NoNewLine -ScreenOnly
								Write-logging -logtext ("{0,20}{1}" -f $skuid,' can be removed') -fg Green -ScreenOnly
								[array]$SafeRedundantServices += $DirectSvc
							}
						}
					}
					""
					"Result: RemoveSku: $RemoveSku"

					if (-not($RemoveSku -eq $False)) {
						$Timer.Stop()
						Write-Logging -logtext ""
						Write-Logging -logtext "Continue removing $skuid`? (Default is No)"
						$Readhost = Read-Host "( y / n )" 
						$Timer.Start()
						Switch ($ReadHost) { 
							Y { $RemoveDirect=$True } 
							J { $RemoveDirect=$True } 
							N { $RemoveDirect=$False } 
							Default { $RemoveDirect=$False } 
						}
						if ($RemoveDirect -eq $True) {
							Write-Logging -logtext ""
							Write-Logging -logtext "Removing $skuid from $UPN" -fg Cyan -bg Black
							Try {
								Set-MsolUserLicense -ObjectId $($User.ObjectID) -RemoveLicenses $skuid -ErrorAction Stop
							}
							Catch {
								$ErrorMessage=$_.Exception.Message
								Write-Logging -logtext "$ErrorMessage" -fg Red
							}
						}

						if ($RemoveDirect -eq $False) {
							Write-Logging -logtext ""
							Write-Logging -logtext "Skipping removal of $skuid on $UPN" -fg DarkGray
							Add-Content "$logPath\Skipped.csv" "$UPN;$skuid;Direct;$($directServices -join ', ');$($CriticalServices -join ', ');$($CriticalRedundantServices -join ', ');$($SafeServices -join ', ');$($SafeRedundantServices -join ', ')"					
						}
					}

					if ($RemoveSku -eq $False) {
						$Timer.Stop()
						Write-Logging -logtext ""
						Write-Logging -logtext "Would you like to override above and continue removing $skuid`? (Default is No)"
						$Readhost = Read-Host "( y / n )" 
						$Timer.Start()
						Switch ($ReadHost) { 
							Y { $OverrideRemoveDirect=$True } 
							J { $OverrideRemoveDirect=$True } 
							N { $OverrideRemoveDirect=$False } 
							Default { $OverrideRemoveDirect=$False } 
						}
						if ($OverrideRemoveDirect -eq $True) {
							Write-Logging -logtext ""
							Write-Logging -logtext "Removing $skuid from $UPN" -fg Cyan -bg Black
							Try {
								Set-MsolUserLicense -ObjectId $($User.ObjectID) -RemoveLicenses $skuid -ErrorAction Stop
							}
							Catch {
								$ErrorMessage=$_.Exception.Message
								Write-Logging -logtext "$ErrorMessage" -fg Red
							}
						}

						if ($OverrideRemoveDirect -eq $False) {
							Write-Logging -logtext ""
							Write-Logging -logtext "Skipping removal of $skuid on $UPN (Removal of license would cause critical service loss)" -fg DarkGray
							Add-Content "$logPath\Skipped.csv" "$UPN;$skuid;Direct;$($directServices -join ', ');$($CriticalServices -join ', ');$($CriticalRedundantServices -join ', ');$($SafeServices -join ', ');$($SafeRedundantServices -join ', ')"					
						}
					}
				}
			} # end direct services validation

			# if the ratio between Directly licensed services and Group licensed services are 1:1, then proceed with removing the Directly licensed SKU
			if ("Direct+Group" -in $ServiceValidationPerSku[$skuid].Values) {
				if (-not("Extra (Direct)" -in $ServiceValidationPerSku[$skuid].Values)) {
					Write-Logging -logtext ""
					Write-Logging -logtext "Removing $skuid from $UPN" -fg Green -bg Black
					$Timer.Stop()
					$Confirm=Read-Host "Confirm (y/n)"
					# kept for debugging:
					#$Confirm="y"
					$Timer.Start()
					if ($Confirm -ieq "y") {
						Try {
							Set-MsolUserLicense -ObjectId $($User.ObjectID) -RemoveLicenses $skuid -ErrorAction Stop
						}
						Catch {
							$ErrorMessage=$_.Exception.Message
							Write-Logging -logtext "$ErrorMessage" -fg Red
						}
					}
				}
			}
		} # end remove licenses per sku
	} # end per license / skupack loop
} # end if $User
$Timer.Stop()
Write-Logging -logtext ""
Write-Logging -logtext "Script took $($Timer.Elapsed.Seconds),$($Timer.Elapsed.MilliSeconds) seconds to run." -fg Yellow