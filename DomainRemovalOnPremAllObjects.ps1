$DomainToReplace = Read-Host -Prompt 'What Domain to check/replace? (e.g. contoso.com)?'
$DomainToUse = Read-Host -Prompt 'What domain to use (e.g. contoso.onmicrosoft.com)?'
$LoggingOnly = Read-Host -Prompt 'Type "Log" for Logging only or "Remove" for Logging AND removal of domain email address'

$DomaintoReplace2 = '*@' + $DomainToReplace

$OUToFilterBy1 = Read-Host -Prompt 'What top level OU should we look in for recipients? (e.g. contoso.com/ContosoUsers) '
$OUToFilterBy2 = Read-Host -Prompt 'What top level OU should we look in for users? (e.g. OU=ContosoUsers,dc=contoso,dc=com) '

$EmailRegex = '^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$'

$Properties = 	@(	'name',	'msExchRecipientDisplayType',	'msExchRecipientTypeDetails',	'ObjectGUID',	'UserPrincipalName',	'ProxyAddresses',	'DistinguishedName', 'Mail'	)
$GUIDs = New-Object System.Collections.Generic.List[System.Object]

$ADUserList = Get-ADUser -SearchBase $OUToFilterBy2 -filter * -resultsetsize 100000 #-properties $Properties | Select $Properties 
ForEach ($AdUser in $AdUserList) {
	$GUIDs.Add($AdUser.ObjectGUID.ToString())
	}
Clear-Variable AdUserList

$ADGroupList = Get-ADGroup -SearchBase $OUToFilterBy2 -filter * -resultsetsize 100000 #-properties $Properties | Select $Properties 
ForEach ($ADGroup in $ADGroupList) {
	$GUIDs.Add($ADGroup.ObjectGUID.ToString())
	}
clear-Variable ADGroupList

$ADContactList = Get-ADObject -SearchBase $OUToFilterBy2 -filter 'objectclass -eq "Contact"' -resultsetsize 100000  # -properties $Properties | Select $Properties 
ForEach ($ADContact in $ADContactList) {
	$GUIDs.Add($ADContact.ObjectGUID.ToString())
	}	
clear-Variable ADContactList

$report=@()

#Let's update UPNs first**************
Write-Host "Let's update UPNs" -ForegroundColor Green
ForEach ($GUID in $GUIDs) {

$Object = Get-ADObject $GUID -properties $Properties | Select $Properties

			If		($Object.msExchRecipientDisplayType -eq (-2147483642)) { $msExchObjectType = "Remote User Mailbox" }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (-2147481850)) { $msExchObjectType = "Remote Room Mailbox" }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (-2147481594)) { $msExchObjectType = "Remote Equipment Mailbox" }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (0)) 		   { $msExchObjectType = "Shared Mailbox" }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (1))           { $msExchObjectType = "Distribution group" }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (6) -and ($Object.msExchRecipientTypeDetails -eq (128)))           { $msExchObjectType = "Mail user" }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (7))           { $msExchObjectType = "Room mailbox" }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (8))           { $msExchObjectType = "Equipment mailbox" }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (1073741824))  { $msExchObjectType = "User mailbox" }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (1073741833))  { $msExchObjectType = "Mail-enabled Security group" }
             else { $msExchObjectType = "Non-mail enabled object"}
			
		$ThisUPN = $Object.userprincipalname
	IF ($ThisUPN -ne $NULL) 
		{
		Write-Host "Checking " $Object.Name
		$ThisUPN = $Object.UserPrincipalName
		$ThisUPNDomain = $ThisUPN.ToString().Split("@")[1]
		$ThisUPNUserNamePortion = $ThisUPN.ToString().Split("@")[0]
		#Is the mailbox's UPN matching our domain to remove?
		If ($ThisUPNDomain -eq $DomainToReplace) {
			$UserObj = New-Object PSObject
			$UserObj | Add-Member -membertype Noteproperty -Name "Time" -value (get-date).ToString("yyyyMMdd-HH:mm:ss.fff")
			$UserObj | Add-Member -membertype NoteProperty -Name "GUID" -Value $Object.ObjectGUID
			$UserObj | Add-Member -membertype NoteProperty -Name "Name" -Value $Object.Name
			$UserObj | Add-Member -Membertype NoteProperty -Name "RecipientType" -Value $msExchObjectType
			Write-host "UPN domain matches! Let's replace it"
			#Update the UPN to the replacement
			$NewAddressToUse = $ThisUPNUserNamePortion + "@" + $DomainToUse
			If ($LoggingOnly -eq "Remove"){
			$Command = 'set-adUser -identity $Object.ObjectGUID -userprincipalname $NewAddressToUse'
			iex $Command -WarningVariable Warning1 -ErrorVariable Error1
			$WarningOutput = $($Warning1 -join [Environment]::NewLine)
			$ErrorOutput = $($Error1 -join [Environment]::NewLine)
			$UserObj | Add-Member -Membertype NoteProperty -Name "Warning" -Value $WarningOutput
			$UserObj | Add-Member -Membertype NoteProperty -Name "Error" -Value $ErrorOutput
			$Actionvalue = "UPN changed from " + $ThisUPN + " to " + $NewAddressToUse
			} else {
			$Actionvalue = "UPN needs to be changed from " + $ThisUPN + " to " + $NewAddressToUse
			}
			Write-Host $Actionvalue
			$UserObj | Add-Member -Membertype NoteProperty -Name "Action" -Value $Actionvalue
			$report += $UserObj
			}
		}
}

Write-host "Finished updating UPNs" -ForegroundColor green
#Let's take a breather
Start-Sleep -milliseconds 1000

#Let's update the Mail attribute next****************
Write-Host "Let's update the mail attribute next****************"
ForEach ($GUID in $GUIDs){
$Object = Get-ADObject $GUID -properties $Properties | Select $Properties

		If		($Object.msExchRecipientDisplayType -eq (-2147483642)) { 
			$msExchObjectType = "Remote User Mailbox" 
			$ObjectInScope = 1 }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (-2147481850)) { 
			$msExchObjectType = "Remote Room Mailbox" 
			$ObjectInScope = 1 }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (-2147481594)) { 
			$msExchObjectType = "Remote Equipment Mailbox"  
			$ObjectInScope = 1 }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (0)) 		   { 
			$msExchObjectType = "Shared Mailbox"  
			$ObjectInScope = 1 }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (1))           { 
			$msExchObjectType = "Distribution group" 
			$ObjectInScope = 1 }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (6) -and ($Object.msExchRecipientTypeDetails -eq (128)))           { 
			$msExchObjectType = "Mail user"  
			$ObjectInScope = 1}
			ElseiF  ($Object.msExchRecipientDisplayType -eq (7))           { 
			$msExchObjectType = "Room mailbox" 
			$ObjectInScope = 1 }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (8))           { 
			$msExchObjectType = "Equipment mailbox"  
			$ObjectInScope = 1}
			ElseiF  ($Object.msExchRecipientDisplayType -eq (1073741824))  { 
			$msExchObjectType = "User mailbox"  
			$ObjectInScope = 1}
			ElseiF  ($Object.msExchRecipientDisplayType -eq (1073741833))  { 
			$msExchObjectType = "Mail-enabled Security group" 
			$ObjectInScope = 1 }
		else { 
			 $msExchObjectType = "Non-mail enabled object"  
			 $ObjectInScope = 0}

	#If ($ObjectInScope -eq 1) {
		Write-Host "Checking mail attribute on " $Object.Name
		$ThisMailAttrib = $Object.Mail
		$DidItMatch = $ThisMailAttrib -match $EmailRegex
		if ($DidItMatch) {
			$ThisMailAttribSMTPDomain = $ThisMailAttrib.ToString().Split("@")[1]
			$ThisMailAttribUserNamePortion = $ThisMailAttrib.ToString().Split("@")[0]
			If ($ThisMailAttribSMTPDomain -eq $DomainToReplace) {
				$UserObj = New-Object PSObject
				$UserObj | Add-Member -membertype Noteproperty -Name "Time" -value (get-date).ToString("yyyyMMdd-HH:mm:ss.fff")
				$UserObj | Add-Member -membertype NoteProperty -Name "GUID" -Value $Object.ObjectGUID
				$UserObj | Add-Member -membertype NoteProperty -Name "name" -Value $Object.name
				$UserObj | Add-Member -Membertype NoteProperty -Name "RecipientType" -Value $msExchObjectType
				Write-Host "Current Mail attribute is " $ThisMailAttrib
				$NewAddressToUse = $ThisMailAttribUserNamePortion + "@" + $DomainToUse
				If ($LoggingOnly -eq "Remove"){
					$Command = 'Set-ADObject $Object.ObjectGUID -replace @{Mail = $NewAddressToUse } '
					iex $Command -WarningVariable Warning1 -ErrorVariable Error1
					$WarningOutput = $($Warning1 -join [Environment]::NewLine)
					$ErrorOutput = $($Error1 -join [Environment]::NewLine)
				$UserObj | Add-Member -Membertype NoteProperty -Name "Warning" -Value $WarningOutput
				$UserObj | Add-Member -Membertype NoteProperty -Name "Error" -Value $ErrorOutput
				$Actionvalue = "Mail attribute updated to " + $NewAddressToUse
				} else {
				$Actionvalue = "Mail attribute needs to change to " + $NewAddressToUse
				}
				Write-Host $Actionvalue
				$UserObj | Add-Member -Membertype NoteProperty -Name "Action" -Value $Actionvalue
				$report += $UserObj
				}
			}
		#}
}


#Let's update primary SMTP addresses next****************
Write-Host "Let's update primary SMTP addresses" -ForegroundColor Yellow
ForEach ($GUID in $GUIDs){
	
$Object = Get-ADObject $GUID -properties $Properties | Select $Properties

			If		($Object.msExchRecipientDisplayType -eq (-2147483642)) { 
			$msExchObjectType = "Remote User Mailbox" 
			$ObjectInScope = 1 }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (-2147481850)) { 
			$msExchObjectType = "Remote Room Mailbox" 
			$ObjectInScope = 1 }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (-2147481594)) { 
			$msExchObjectType = "Remote Equipment Mailbox"  
			$ObjectInScope = 1 }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (0)) 		   { 
			$msExchObjectType = "Shared Mailbox"  
			$ObjectInScope = 1 }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (1))           { 
			$msExchObjectType = "Distribution group" 
			$ObjectInScope = 1 }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (6) -and ($Object.msExchRecipientTypeDetails -eq (128)))           { 
			$msExchObjectType = "Mail user"  
			$ObjectInScope = 1}
			ElseiF  ($Object.msExchRecipientDisplayType -eq (7))           { 
			$msExchObjectType = "Room mailbox" 
			$ObjectInScope = 1 }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (8))           { 
			$msExchObjectType = "Equipment mailbox"  
			$ObjectInScope = 1}
			ElseiF  ($Object.msExchRecipientDisplayType -eq (1073741824))  { 
			$msExchObjectType = "User mailbox"  
			$ObjectInScope = 1}
			ElseiF  ($Object.msExchRecipientDisplayType -eq (1073741833))  { 
			$msExchObjectType = "Mail-enabled Security group" 
			$ObjectInScope = 1 }
             else { 
			 $msExchObjectType = "Non-mail enabled object"  
			 $ObjectInScope = 0}
			
#	If ($ObjectInScope -eq 1) {
			Write-Host "Checking primary SMTP on " $Object.Name
            
			$ThisPrimarySMTP = Get-ADObject -identity $Object.ObjectGUID -Properties ProxyAddresses | Select -ExpandProperty ProxyAddresses | ? {$_ -clike "SMTP:*"}
            if ($ThisPrimarySMTP) {
			$ThisPrimarySMTPAddressOnly = $ThisPrimarySMTP.ToString().Split(":")[1]
			$ThisPrimarySMTPDomain = $ThisPrimarySMTPAddressOnly.ToString().Split("@")[1]
			$ThisPrimarySMTPUserNamePortion = $ThisPrimarySMTPAddressOnly.ToString().Split("@")[0]
			If ($ThisPrimarySMTPDomain -eq $DomainToReplace) {
				$UserObj = New-Object PSObject
				$UserObj | Add-Member -membertype Noteproperty -Name "Time" -value (get-date).ToString("yyyyMMdd-HH:mm:ss.fff")
				$UserObj | Add-Member -membertype NoteProperty -Name "GUID" -Value $Object.ObjectGUID
				$UserObj | Add-Member -membertype NoteProperty -Name "name" -Value $Object.name
				$UserObj | Add-Member -Membertype NoteProperty -Name "RecipientType" -Value $msExchObjectType
				Write-Host "Current PrimarySMTP is " $ThisPrimarySMTPAddressOnly
				$NewAddressToUse = $ThisPrimarySMTPUserNamePortion + "@" + $DomainToUse
				If ($LoggingOnly -eq "Remove"){
					Set-ADObject -identity $Object.ObjectGUID -Add @{ProxyAddresses = "smtp:" + $NewAddressToUse}
					Get-ADObject -identity $Object.ObjectGUID -properties Mail,ProxyAddresses |
						ForEach {
							$proxies = $_.ProxyAddresses |
							ForEach-Object{
								$a = $_ -replace 'SMTP','smtp'
								if($a -match $NewAddressToUse){
								$a -replace 'smtp','SMTP'
								}else{
								$a
								}
							}
						$_.ProxyAddresses = $proxies
						$Command = 'Set-ADObject -instance $_'
						iex $Command -WarningVariable Warning1 -ErrorVariable Error1
						$WarningOutput = $($Warning1 -join [Environment]::NewLine)
						$ErrorOutput = $($Error1 -join [Environment]::NewLine)
					}
				$UserObj | Add-Member -Membertype NoteProperty -Name "Warning" -Value $WarningOutput
				$UserObj | Add-Member -Membertype NoteProperty -Name "Error" -Value $ErrorOutput
				$Actionvalue = "Primary SMTP updated to " + $NewAddressToUse
				} else {
				$Actionvalue = "Primary SMTP needs to change to " + $NewAddressToUse
				}
				Write-Host $Actionvalue
				$UserObj | Add-Member -Membertype NoteProperty -Name "Action" -Value $Actionvalue
				$report += $UserObj
            }
			}
	#}
}


#Let's pause for 3 seconds
Start-sleep -seconds 3

#Let's remove all email addresses that are using the pertinent domain
Write-Host "Let's remove the addresses" -ForegroundColor magenta
ForEach ($GUID in $GUIDs) {
		
	$Object = Get-ADObject $GUID -properties $Properties | Select $Properties
						If		($Object.msExchRecipientDisplayType -eq (-2147483642)) { 
			$msExchObjectType = "Remote User Mailbox" 
			$ObjectInScope = 1 }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (-2147481850)) { 
			$msExchObjectType = "Remote Room Mailbox" 
			$ObjectInScope = 1 }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (-2147481594)) { 
			$msExchObjectType = "Remote Equipment Mailbox"  
			$ObjectInScope = 1 }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (0)) 		   { 
			$msExchObjectType = "Shared Mailbox"  
			$ObjectInScope = 1 }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (1))           { 
			$msExchObjectType = "Distribution group" 
			$ObjectInScope = 1 }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (6) -and ($Object.msExchRecipientTypeDetails -eq (128)))           { 
			$msExchObjectType = "Mail user"  
			$ObjectInScope = 1}
			ElseiF  ($Object.msExchRecipientDisplayType -eq (7))           { 
			$msExchObjectType = "Room mailbox" 
			$ObjectInScope = 1 }
			ElseiF  ($Object.msExchRecipientDisplayType -eq (8))           { 
			$msExchObjectType = "Equipment mailbox"  
			$ObjectInScope = 1}
			ElseiF  ($Object.msExchRecipientDisplayType -eq (1073741824))  { 
			$msExchObjectType = "User mailbox"  
			$ObjectInScope = 1}
			ElseiF  ($Object.msExchRecipientDisplayType -eq (1073741833))  { 
			$msExchObjectType = "Mail-enabled Security group" 
			$ObjectInScope = 1 }
			else { 
			$msExchObjectType = "Non-mail enabled object"  
			$ObjectInScope = 0}
			
#	If ($ObjectInScope -eq 1) { 
		Write-Host "Checking addresses on " $Object.Name
		Foreach ($Address in $Object.ProxyAddresses){
			Write-Host "Checking " $Address
			$AddressDomain = (($Address.ToString().Split(":")[1]).Split("@")[1])	   # For 'SMTP:username@contoso.com' - "contoso.com"
			If ($AddressDomain -eq $DomainToReplace) {
			$UserObj = New-Object PSObject
			$UserObj | Add-Member -membertype Noteproperty -Name "Time" -value (get-date).ToString("yyyyMMdd-HH:mm:ss.fff")
			$UserObj | Add-Member -membertype NoteProperty -Name "GUID" -Value $Object.ObjectGUID
			$UserObj | Add-Member -membertype NoteProperty -Name "name" -Value $Object.name
			$UserObj | Add-Member -Membertype NoteProperty -Name "RecipientType" -Value $msExchObjectType
			#Remove the address
			Write-Host "Removing the address from the mailbox..."
			#Start-sleep -seconds 1
			If ($LoggingOnly -eq "Remove"){
				$Command = 'Set-ADObject $Object.ObjectGUID -Remove  @{ProxyAddresses= $Address}'
				iex $Command -WarningVariable Warning1 -ErrorVariable Error1
				$WarningOutput = $($Warning1 -join [Environment]::NewLine)
				$ErrorOutput = $($Error1 -join [Environment]::NewLine)
				$UserObj | Add-Member -Membertype NoteProperty -Name "Warning" -Value $WarningOutput
				$UserObj | Add-Member -Membertype NoteProperty -Name "Error" -Value $ErrorOutput
				$Actionvalue = "Email address removed - " + $Address + " from " + $Object.name
				} else {
				$Actionvalue = "Email address removal needed - " + $Address + " from " + $Object.name
				}
			Write-Host $Actionvalue -foregroundcolor DarkRed
			$UserObj | Add-Member -Membertype NoteProperty -Name "Action" -Value $Actionvalue 
			$report += $UserObj
			} else {
			Write-Host "This address does not need to be removed"
			}
		#Start-sleep -seconds 1
		}		
	#}
}

Write-Host "All relevant email addresses removed" -ForegroundColor magenta

If ($LoggingOnly -eq "Remove"){
$FileName = "C:\powershell\DomainCutover\DomainRemovalReport-OnPREM-" + $DomaintoReplace + "-" + $((get-date).ToString("yyyyMMdd-HHmm")) + "-Remove.csv"
} else {
$FileName = "C:\powershell\DomainCutover\DomainRemovalReport-OnPREM-" + $DomaintoReplace + "-" + $((get-date).ToString("yyyyMMdd-HHmm")) + "-LoggingOnly.csv"
}

$Report | Export-CSV $FileName -NoType 

Write-Host "Script execution completed.  Log file - " $FileName