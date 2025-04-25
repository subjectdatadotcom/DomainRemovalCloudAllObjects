<#
.SYNOPSIS
This script automates the management of email addresses and domains for Exchange Online recipients, ensuring compliance and domain consistency across user mailboxes and groups.

.DESCRIPTION
The script checks for and installs necessary PowerShell modules (ExchangeOnlineManagement and MSOnline), connects to Exchange Online, and imports domains from a CSV file for processing. It updates email properties like UserPrincipalName and primary SMTP addresses, removes outdated email addresses, and logs these actions for each recipient based on their type (UserMailbox, MailUser, GroupMailbox, etc.). Additionally, it identifies accounts that require manual intervention.

.NOTES
The script requires administrative credentials for Exchange Online and is designed for environments that need regular updates to recipient email configurations to maintain domain integrity and compliance.

.AUTHOR
SubjectData

.EXAMPLE
.\DomainRemovalCloudAllObjects.ps1
Executes the script, processing email addresses according to the domains listed in 'Domains.csv', updating properties, removing obsolete addresses, and generating a report of actions taken.
#>


# Ensure ExchangeOnlineManagement module is installed and imported
$exchangeModule = "ExchangeOnlineManagement"

if (-not (Get-Module -Name $exchangeModule -ListAvailable)) {
    Write-Host "$exchangeModule module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name $exchangeModule -Force -Scope CurrentUser
}

Import-Module $exchangeModule -Force
Write-Host "$exchangeModule module successfully loaded." -ForegroundColor Green

# Ensure MSOnline module is installed and imported
$msolModule = "MSOnline"

if (-not (Get-Module -Name $msolModule -ListAvailable)) {
    Write-Host "$msolModule module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name $msolModule -Force -Scope CurrentUser
}

Import-Module $msolModule -Force
Write-Host "$msolModule module successfully loaded." -ForegroundColor Green


# Connect to Exchange Online
try {
    Connect-ExchangeOnline 
    Connect-MsolService
} catch {
    Write-Host "Failed to connect to Exchange Online. Please check your credentials and try again." -ForegroundColor Red
    exit
}

# Get the directory of the current script
$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Define the location of the CSV file containing OneDrive user emails
$XLloc = "$myDir\"

try {
    # Import the list of OneDrive users from the CSV file
    $Domains = import-csv ($XLloc + "Domains.csv").ToString() | Select-Object -ExpandProperty Domain
} catch {
    # Handle the error if the CSV file is not found
    Write-Host "No CSV file to read" -BackgroundColor Black -ForegroundColor Red
    exit
}


$DomainToUse = Read-Host -Prompt 'What domain to use (e.g. contoso.onmicrosoft.com)?'
$DomainToUse = $DomainToUse.Trim()
#$DomainToUse = "M365x76832558.onmicrosoft.com"#$DomainToUse.Trim()

# Prompt for logging option
$LoggingOnly = Read-Host -Prompt 'Type "Log" for Logging only or "Remove" for Logging AND removal of domain email address'
$LoggingOnly = $LoggingOnly.Trim()
#$LoggingOnly = "Remove"

# Initialize report array
$report = @()

foreach ($DomainToReplace in $Domains) {
    $DomainToReplace2 = "*@" + $DomainToReplace
    Write-Host "Processing domain: $DomainToReplace" -ForegroundColor Cyan

    # Get recipients with matching email domain (licensed accounts with mailboxes, groups, etc)
    $RecipientList = Get-EXORecipient -ResultSize unlimited | Where-Object { $_.EmailAddresses -like $DomainToReplace2 }

    foreach ($Recipient in $RecipientList) {
        $GUID = $Recipient.ExternalDirectoryObjectID.ToString()
        $PrimarySMTP = $Recipient.PrimarySMTPAddress
        $DisplayName = $Recipient.DisplayName
        $RecipientType = $Recipient.RecipientType
        $RecipientDetails = $Recipient.RecipientTypeDetails

        $UserObj = New-Object PSObject
        $UserObj | Add-Member -MemberType NoteProperty -Name "Time" -Value (Get-Date).ToString("yyyyMMdd-HH:mm:ss.fff")
        $UserObj | Add-Member -MemberType NoteProperty -Name "GUID" -Value $GUID
        $UserObj | Add-Member -MemberType NoteProperty -Name "PrimaryEmailAddress" -Value $PrimarySMTP
        $UserObj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $DisplayName
        $UserObj | Add-Member -MemberType NoteProperty -Name "RecipientType" -Value $RecipientType
        $UserObj | Add-Member -MemberType NoteProperty -Name "RecipientTypeDetails" -Value $RecipientDetails
                
        # Handling UserMailbox - Completed - Green
        if ($RecipientDetails -eq "UserMailbox") {
            $ThisMailbox = Get-Mailbox -Identity $GUID
            $UserObj | Add-Member -MemberType NoteProperty -Name "IsObjectSyncedAADC" -Value $ThisMailbox.IsDirSynced

            # Initialize Warning and Error logs as empty arrays
            $WarningLog = @()
            $ErrorLog = @()

             Write-Host "User Mailbox - checking " $ThisMailUser.DisplayName -ForegroundColor Green

            try {
                # --- Update UPN ---
                $ThisUPN = $ThisMailbox.UserPrincipalName
                $ThisUPNDomain = $ThisUPN.Split("@")[1]
                $ThisUPNUserNamePortion = $ThisUPN.Split("@")[0]

                if ($ThisUPNDomain -eq $DomainToReplace) {
                    $NewAddressToUse = "$ThisUPNUserNamePortion@$DomainToUse"
                    
                    if ($LoggingOnly -eq "Remove") {
                        $Command = "Get-MSolUser -UserPrincipalName $ThisUPN | Set-MsolUserPrincipalName -NewUserPrincipalName $NewAddressToUse"
                        iex $Command -WarningVariable Warning1 -ErrorVariable Error1
                        Start-Sleep -Seconds 1  # Allow the update to take effect

                        # Append Warning and Error messages with UPN-specific tags
                        if ($Warning1) { $WarningLog += "UPN Warning: " + ($Warning1 -join " | ") }
                        if ($Error1) { $ErrorLog += "UPN Error: " + ($Error1 -join " | ") }
                    }

                    $UserObj | Add-Member -MemberType NoteProperty -Name "Action_UPN" -Value "UPN changed from $ThisUPN to $NewAddressToUse"
                    
                }

                # --- Update Primary SMTP ---
                $ThisPrimarySMTP = $ThisMailbox.PrimarySMTPAddress
                $ThisPrimarySMTPDomain = $ThisPrimarySMTP.Split("@")[1]
                $ThisPrimarySMTPUserNamePortion = $ThisPrimarySMTP.Split("@")[0]

                if ($ThisPrimarySMTPDomain -eq $DomainToReplace) {
                    $NewSMTP = "$ThisPrimarySMTPUserNamePortion@$DomainToUse"
                    
                    if ($LoggingOnly -eq "Remove") {
                        $Command = "Set-Mailbox $GUID -WindowsEmailAddress $NewSMTP"
                        iex $Command -WarningVariable Warning2 -ErrorVariable Error2
                        Start-Sleep -Seconds 1  # Allow update to apply

                        # Append Warning and Error messages with SMTP-specific tags
                        if ($Warning2) { $WarningLog += "SMTP Warning: " + ($Warning2 -join " | ") }
                        if ($Error2) { $ErrorLog += "SMTP Error: " + ($Error2 -join " | ") }
                    }

                    $UserObj | Add-Member -MemberType NoteProperty -Name "Action_SMTP" -Value "Primary SMTP changed from $ThisPrimarySMTP to $NewSMTP"
                    
                }

                # --- Remove Email Addresses matching the domain ---
                Write-Host "Checking User Mailbox addresses on " $ThisMailUser.DisplayName
                foreach ($Address in $ThisMailbox.EmailAddresses) {
                    $AddressDomain = (($Address.Split(":")[1]).Split("@")[1])

                    if ($AddressDomain -eq $DomainToReplace) {
                        
                        if ($LoggingOnly -eq "Remove") {
                            $Command = "Set-Mailbox $GUID -EmailAddresses @{Remove='$Address'}"
                            iex $Command -WarningVariable Warning3 -ErrorVariable Error3
                            Start-Sleep -Seconds 1  # Allow update to apply

                            # Append Warning and Error messages with Email Removal-specific tags
                            if ($Warning3) { $WarningLog += "Email Removal Warning: " + ($Warning3 -join " | ") }
                            if ($Error3) { $ErrorLog += "Email Removal Error: " + ($Error3 -join " | ") }
                        }

                        $UserObj | Add-Member -MemberType NoteProperty -Name "Action_EmailRemoval" -Value "Email address removed - $Address from $ThisMailbox.DisplayName"
                        
                    }
                }

                # -- Append Exception Type & Message for each condition --
                if ($WarningLog -or $ErrorLog) {
                    $UserObj | Add-Member -MemberType NoteProperty -Name "ExceptionType" -Value "Warnings & Errors"
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Message" -Value (($WarningLog + $ErrorLog) -join " | ")
                }
            }
            catch {
                Write-Host $ThisMailbox.UserPrincipalName -ForegroundColor DarkMagenta
                # Capture General Errors
                $ErrorLog += "General Error: $($_.Exception.Message)"                
                $UserObj | Add-Member -MemberType NoteProperty -Name "GeneralError" -Value $ErrorLog
            }
        }

        # Handling Shared Mailbox
       elseif ($RecipientDetails -eq "SharedMailbox") {
            $ThisMailbox = Get-Mailbox -Identity $GUID
            $UserObj | Add-Member -MemberType NoteProperty -Name "IsObjectSyncedAADC" -Value $ThisMailbox.IsDirSynced

            $WarningLog = @()
            $ErrorLog = @()

            Write-Host "SharedMailbox - checking " $ThisMailbox.DisplayName -ForegroundColor Green

            try {
                # Prepare UPN Variables
                $CurrentUPN = $ThisMailbox.UserPrincipalName
                $NewUPN = $ThisMailbox.WindowsEmailAddress

                # --- UPN Update (First) ---
                $ThisUPN = $ThisMailbox.UserPrincipalName
                $ThisUPNDomain = $ThisUPN.Split("@")[1]
                $ThisUPNUserNamePortion = $ThisUPN.Split("@")[0]

                if ($ThisUPNDomain -eq $DomainToReplace) {
                    $NewAddressToUse = "$ThisUPNUserNamePortion@$DomainToUse"

                    if ($LoggingOnly -eq "Remove") {
                        try {
                            $Command = "Get-MsolUser -UserPrincipalName `"$ThisUPN`" | Set-MsolUserPrincipalName -NewUserPrincipalName `"$NewAddressToUse`""
                            iex $Command -WarningVariable WarningUPN -ErrorVariable ErrorUPN
                            Start-Sleep -Seconds 1  # Allow update to apply

                            # Append Warning and Error messages
                            if ($WarningUPN) { $WarningLog += "UPN Warning: " + ($WarningUPN -join " | ") }
                            if ($ErrorUPN) { $ErrorLog += "UPN Error: " + ($ErrorUPN -join " | ") }

                            Write-Host "UPN changed from $ThisUPN to $NewAddressToUse for SharedMailbox" -ForegroundColor Green
                            $UserObj | Add-Member -MemberType NoteProperty -Name "Action_UPN" -Value "UPN changed from $ThisUPN to $NewAddressToUse"
                        } catch {
                            Write-Warning "Failed to change UPN from $ThisUPN to $($NewAddressToUse): $($_.Exception.Message)"
                            $ErrorLog += "UPN Error: $($_.Exception.Message)"
                        }
                    }
                }


                # --- Primary SMTP Update ---
                $ThisPrimarySMTP = $ThisMailbox.PrimarySMTPAddress
                $ThisPrimarySMTPDomain = $ThisPrimarySMTP.Split("@")[1]
                $ThisPrimarySMTPUserNamePortion = $ThisPrimarySMTP.Split("@")[0]

                if ($ThisPrimarySMTPDomain -eq $DomainToReplace) {
                    $NewSMTP = "$ThisPrimarySMTPUserNamePortion@$DomainToUse"

                    if ($LoggingOnly -eq "Remove") {
                        try {
                            $Command = "Set-Mailbox -Identity `"$GUID`" -WindowsEmailAddress `"$NewSMTP`""
                            iex $Command -WarningVariable WarningSMTP -ErrorVariable ErrorSMTP
                            Start-Sleep -Seconds 1

                            if ($WarningSMTP) { $WarningLog += "SMTP Warning: " + ($WarningSMTP -join " | ") }
                            if ($ErrorSMTP) { $ErrorLog += "SMTP Error: " + ($ErrorSMTP -join " | ") }

                            Write-Host "Primary SMTP changed from $ThisPrimarySMTP to $NewSMTP" -ForegroundColor Green
                            $UserObj | Add-Member -MemberType NoteProperty -Name "Action_SMTP" -Value "Primary SMTP changed from $ThisPrimarySMTP to $NewSMTP"
                        } catch {
                            Write-Warning "Failed to change Primary SMTP: $($_.Exception.Message)"
                            $ErrorLog += "SMTP Error: $($_.Exception.Message)"
                        }
                    }
                }

                # --- Remove Email Addresses matching the domain ---
                Write-Host "Checking SharedMailbox addresses on " $ThisMailbox.DisplayName
                foreach ($Address in $ThisMailbox.EmailAddresses) {
                    $AddressValue = $Address -split ":" | Select-Object -Last 1
                    $AddressDomain = $AddressValue.Split("@")[1]

                    if ($AddressDomain -eq $DomainToReplace) {
                        if ($LoggingOnly -eq "Remove") {
                            try {
                                $Command = "Set-Mailbox -Identity `"$GUID`" -EmailAddresses @{Remove='$Address'}"
                                iex $Command -WarningVariable WarningRemove -ErrorVariable ErrorRemove
                                Start-Sleep -Seconds 1

                                if ($WarningRemove) { $WarningLog += "Email Removal Warning: " + ($WarningRemove -join " | ") }
                                if ($ErrorRemove) { $ErrorLog += "Email Removal Error: " + ($ErrorRemove -join " | ") }

                                Write-Host "Removed $Address from $ThisMailbox.DisplayName" -ForegroundColor DarkRed
                                $UserObj | Add-Member -MemberType NoteProperty -Name "Action_EmailRemoval" -Value "Email address removed - $Address from $ThisMailbox.DisplayName"
                            } catch {
                                Write-Warning "Failed to remove email address $($Address): $($_.Exception.Message)"
                                $ErrorLog += "Email Removal Error: $($_.Exception.Message)"
                            }
                        }
                    } else {
                        Write-Host "This address does not need to be removed"
                    }
                }

                # -- Append Exception Type & Message for each condition --
                if ($WarningLog -or $ErrorLog) {
                    $UserObj | Add-Member -MemberType NoteProperty -Name "ExceptionType" -Value "Warnings & Errors"
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Message" -Value (($WarningLog + $ErrorLog) -join " | ")
                }
            }
            catch {
                Write-Host $ThisMailbox.UserPrincipalName -ForegroundColor DarkMagenta
                $ErrorLog += "General Error: $($_.Exception.Message)"
                $UserObj | Add-Member -MemberType NoteProperty -Name "GeneralError" -Value $ErrorLog
            }

            $report += $UserObj
        }


        
        # Handling MailUser - Completed - Yellow
        elseif ($RecipientType -eq "MailUser") {
             try {
                $ThisMailUser = Get-MailUser -Identity $GUID
                $UserObj | Add-Member -MemberType NoteProperty -Name "IsObjectSyncedAADC" -Value $ThisMailUser.IsDirSynced

                # Store errors and warnings
                $WarningLog = @()
                $ErrorLog = @()

                Write-Host "Mail user - checking " $ThisMailUser.DisplayName -ForegroundColor Yellow


                # --- Update UPN ---
                $ThisUPN = $ThisMailbox.UserPrincipalName
                $ThisUPNDomain = $ThisUPN.Split("@")[1]
                $ThisUPNUserNamePortion = $ThisUPN.Split("@")[0]

                if ($ThisUPNDomain -eq $DomainToReplace) {
                    $NewAddressToUse = "$ThisUPNUserNamePortion@$DomainToUse"
                    
                    if ($LoggingOnly -eq "Remove") {
                        $Command = "Get-MSolUser -UserPrincipalName $ThisUPN | Set-MsolUserPrincipalName -NewUserPrincipalName $NewAddressToUse"
                        iex $Command -WarningVariable Warning1 -ErrorVariable Error1
                        Start-Sleep -Seconds 1

                        # Append Warning and Error messages
                        if ($Warning1) { $WarningLog += "UPN Warning: " + ($Warning1 -join " | ") }
                        if ($Error1) { $ErrorLog += "UPN Error: " + ($Error1 -join " | ") }
                    }

                    $UserObj | Add-Member -MemberType NoteProperty -Name "Action_UPN" -Value "UPN changed from $ThisUPN to $NewAddressToUse"
                    
                }

                # -- Primary SMTP Update --
                $ThisPrimarySMTP = $ThisMailUser.PrimarySMTPAddress
                $ThisPrimarySMTPDomain = $ThisPrimarySMTP.Split("@")[1]
                $ThisPrimarySMTPUserNamePortion = $ThisPrimarySMTP.Split("@")[0]

                if ($ThisPrimarySMTPDomain -eq $DomainToReplace) {
                    $NewSMTP = "$ThisPrimarySMTPUserNamePortion@$DomainToUse"
                    
                    if ($LoggingOnly -eq "Remove") {
                        $Command = "Set-MailUser $GUID -PrimarySMTPAddress $NewSMTP"
                        iex $Command -WarningVariable Warning2 -ErrorVariable Error2
                        Start-Sleep -Seconds 1

                        # Append Warning and Error messages
                        if ($Warning2) { $WarningLog += "Primary SMTP Warning: " + ($Warning2 -join " | ") }
                        if ($Error2) { $ErrorLog += "Primary SMTP Error: " + ($Error2 -join " | ") }
                    }

                    $UserObj | Add-Member -MemberType NoteProperty -Name "Action_SMTP" -Value "Primary SMTP changed from $ThisPrimarySMTP to $NewSMTP"
                    
                }

                # -- Email Address Removal --
                Write-Host "Checking Mail User addresses on " $ThisMailUser.DisplayName
                foreach ($Address in $ThisMailUser.EmailAddresses) {
                    Write-Host "Checking " $Address
                    $AddressDomain = (($Address.Split(":")[1]).Split("@")[1]) # Extract domain from email address

                    if ($AddressDomain -eq $DomainToReplace) {
                        
                        if ($LoggingOnly -eq "Remove") {
                            $Command = "Set-MailUser $GUID -EmailAddresses @{Remove='$Address'}"
                            iex $Command -WarningVariable Warning3 -ErrorVariable Error3
                            Start-Sleep -Seconds 1

                            # Append Warning and Error messages
                            if ($Warning3) { $WarningLog += "Email Removal Warning: " + ($Warning3 -join " | ") }
                            if ($Error3) { $ErrorLog += "Email Removal Error: " + ($Error3 -join " | ") }
                        }

                        $ActionValue = "Email address removed - $Address from $ThisMailUser.DisplayName"
                        Write-Host $ActionValue -ForegroundColor DarkRed
                        $UserObj | Add-Member -MemberType NoteProperty -Name "Action_EmailRemoval" -Value $ActionValue
                        
                    } else {
                        Write-Host "This address does not need to be removed"
                    }
                }

                # Add warnings and errors to log
                #if ($WarningLog) { $UserObj | Add-Member -MemberType NoteProperty -Name "Warnings" -Value ($WarningLog -join " | ") }
                #if ($ErrorLog) { $UserObj | Add-Member -MemberType NoteProperty -Name "Errors" -Value ($ErrorLog -join " | ") }

                # -- Append Exception Type & Message for each condition --
                if ($WarningLog -or $ErrorLog) {
                    $UserObj | Add-Member -MemberType NoteProperty -Name "ExceptionType" -Value "Warnings & Errors"
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Message" -Value (($WarningLog + $ErrorLog) -join " | ")
                }
            }
            catch {
                $ErrorLog += "General MailUser Error: $($_.Exception.Message)"
                $UserObj | Add-Member -MemberType NoteProperty -Name "GeneralError" -Value $ErrorLog
            }
        }

        # Handling Security Groups - Completed - Magenta
        elseif (($RecipientType -match "MailUniversalDistributionGroup|MailUniversalSecurityGroup") -and ($RecipientDetails -match "MailUniversalDistributionGroup|MailUniversalSecurityGroup")) {
            try {
                # Fetch Distribution/Security Group Information
                $ThisDistroGroup = Get-DistributionGroup -Identity $GUID
                $UserObj | Add-Member -MemberType NoteProperty -Name "IsObjectSyncedAADC" -Value $ThisDistroGroup.IsDirSynced
                #Write-Host "Checking Group - " $ThisDistroGroup.DisplayName

                # Store errors and warnings
                $WarningLog = @()
                $ErrorLog = @()

                Write-Host "MailUniversalDistributionGroup / MailUniversalSecurityGroup User addresses on " $ThisDistroGroup.DisplayName -ForegroundColor Magenta

                $ThisPrimarySMTP = $ThisDistroGroup.PrimarySMTPAddress
                $ThisPrimarySMTPDomain = $ThisPrimarySMTP.Split("@")[1]
                $ThisPrimarySMTPGroupNamePortion = $ThisPrimarySMTP.Split("@")[0]

                # -- Primary SMTP Update --
                if ($ThisPrimarySMTPDomain -eq $DomainToReplace) {
                    $NewSMTP = "$ThisPrimarySMTPGroupNamePortion@$DomainToUse"
                    
                    if ($LoggingOnly -eq "Remove") {
                        $Command = "Set-DistributionGroup $GUID -PrimarySMTPAddress $NewSMTP"
                        iex $Command -WarningVariable Warning1 -ErrorVariable Error1
                        Start-Sleep -Seconds 1

                        # Append Warning and Error messages
                        if ($Warning1) { $WarningLog += "Primary SMTP Warning: " + ($Warning1 -join " | ") }
                        if ($Error1) { $ErrorLog += "Primary SMTP Error: " + ($Error1 -join " | ") }
                    }

                    $UserObj | Add-Member -MemberType NoteProperty -Name "Action_SMTP" -Value "Primary SMTP changed from $ThisPrimarySMTP to $NewSMTP"
                    
                }

                # -- Email Address Removal --
                Write-Host "MailUniversalDistributionGroup / MailUniversalSecurityGroup - Checking addresses on " $ThisDistroGroup.DisplayName
                foreach ($Address in $ThisDistroGroup.EmailAddresses) {
                    Write-Host "Checking " $Address
                    $AddressDomain = (($Address.Split(":")[1]).Split("@")[1]) # Extract domain from email address

                    if ($AddressDomain -eq $DomainToReplace) {
                    
                        if ($LoggingOnly -eq "Remove") {
                            $Command = "Set-DistributionGroup $GUID -EmailAddresses @{Remove='$Address'}"
                            iex $Command -WarningVariable Warning2 -ErrorVariable Error2
                            Start-Sleep -Seconds 1

                            # Append Warning and Error messages
                            if ($Warning2) { $WarningLog += "Email Removal Warning: " + ($Warning2 -join " | ") }
                            if ($Error2) { $ErrorLog += "Email Removal Error: " + ($Error2 -join " | ") }
                        }

                        $ActionValue = "Email address removed - $Address from $ThisDistroGroup.DisplayName"
                        Write-Host $ActionValue -ForegroundColor DarkRed
                        $UserObj | Add-Member -MemberType NoteProperty -Name "Action_EmailRemoval" -Value $ActionValue
                        
                    } else {
                        Write-Host "This address does not need to be removed"
                    }
                }

                # Add warnings and errors to log
                #if ($WarningLog) { $UserObj | Add-Member -MemberType NoteProperty -Name "Warnings" -Value ($WarningLog -join " | ") }
                #if ($ErrorLog) { $UserObj | Add-Member -MemberType NoteProperty -Name "Errors" -Value ($ErrorLog -join " | ") }

                # -- Append Exception Type & Message for each condition --
                if ($WarningLog -or $ErrorLog) {
                    Write-Host "Entered in logs"
                    $UserObj | Add-Member -MemberType NoteProperty -Name "ExceptionType" -Value "Warnings & Errors"
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Message" -Value (($WarningLog + $ErrorLog) -join " | ")
                }
            }
            catch {
                $ErrorLog += "General Group Error: $($_.Exception.Message)"
                $UserObj | Add-Member -MemberType NoteProperty -Name "GeneralError" -Value $ErrorLog
            }
        }


        # Handling MailContact - Completed - Dark Green
        elseif ($RecipientType -eq "MailContact") {
            try {
                # Fetch Mail Contact Information
                Write-Host "Mail Contact found - logging ONLY - " $Recipient.Identity -ForegroundColor DarkGreen
                $ThisMailContact = Get-MailContact -Identity $GUID
                $UserObj | Add-Member -MemberType NoteProperty -Name "IsObjectSyncedAADC" -Value $ThisMailContact.IsDirSynced
        
                # Store errors and warnings
                $WarningLog = @()
                $ErrorLog = @()
        
                $ThisPrimarySMTP = $ThisMailContact.PrimarySMTPAddress
                $ThisPrimarySMTPDomain = $ThisPrimarySMTP.Split("@")[1]
       
                # -- Primary SMTP Logging (No Change) --
                if ($ThisPrimarySMTPDomain -eq $DomainToReplace) {
                    Write-Host "Current Primary SMTP is " $ThisPrimarySMTP
                    $ActionValue = "Primary SMTP Found for Mail Contact - Logging Only - " + $ThisPrimarySMTP
                    Write-Host $ActionValue
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Action_SMTP" -Value $ActionValue
                }
        
                # -- Email Address Logging (No Change) --
                Write-Host "Mail Contact - Checking addresses on " $ThisMailContact.DisplayName
                foreach ($Address in $ThisMailContact.EmailAddresses) {
                    Write-Host "Checking " $Address
                    $AddressDomain = (($Address.Split(":")[1]).Split("@")[1]) # Extract domain from email address
        
                    if ($AddressDomain -eq $DomainToReplace) {
                        $ActionValue = "Email Address Found for Mail Contact - Logging Only - " + $Address
                        Write-Host $ActionValue -ForegroundColor DarkRed
                        $UserObj | Add-Member -MemberType NoteProperty -Name "Action_EmailLogging" -Value $ActionValue
                    } else {
                        Write-Host "This address does not need to be logged"
                    }
                }
        
                # -- Append Exception Type & Message for each condition --
                if ($WarningLog -or $ErrorLog) {
                    $UserObj | Add-Member -MemberType NoteProperty -Name "ExceptionType" -Value "Warnings & Errors"
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Message" -Value (($WarningLog + $ErrorLog) -join " | ")
                }
        
            }
            catch {
                $ErrorLog += "General Mail Contact Error: $($_.Exception.Message)"
                $UserObj | Add-Member -MemberType NoteProperty -Name "GeneralError" -Value $ErrorLog
            }
        }
        
        # Handling DynamicDistributionGroup - Completed - White
        elseif ($RecipientType -eq "DynamicDistributionGroup") {
            try {
                # Fetch Dynamic Distribution Group Information
                $ThisDistroGroup = Get-DynamicDistributionGroup -Identity $GUID
                Write-Host "Checking Dynamic Distribution Group - " $ThisDistroGroup.DisplayName -ForegroundColor White
        
                # Store errors and warnings
                $WarningLog = @()
                $ErrorLog = @()
        
                $ThisPrimarySMTP = $ThisDistroGroup.PrimarySMTPAddress
                $ThisPrimarySMTPDomain = $ThisPrimarySMTP.Split("@")[1]
                $ThisPrimarySMTPGroupNamePortion = $ThisPrimarySMTP.Split("@")[0]
        
                # -- Primary SMTP Update --
                if ($ThisPrimarySMTPDomain -eq $DomainToReplace) {
                    $NewSMTP = "$ThisPrimarySMTPGroupNamePortion@$DomainToUse"
                    
                    if ($LoggingOnly -eq "Remove") {
                        $Command = "Set-DynamicDistributionGroup $GUID -PrimarySMTPAddress $NewSMTP"
                        iex $Command -WarningVariable Warning1 -ErrorVariable Error1
                        Start-Sleep -Seconds 1
        
                        # Append Warning and Error messages
                        if ($Warning1) { $WarningLog += "Primary SMTP Warning: " + ($Warning1 -join " | ") }
                        if ($Error1) { $ErrorLog += "Primary SMTP Error: " + ($Error1 -join " | ") }
                    }
        
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Action_SMTP" -Value "Primary SMTP changed from $ThisPrimarySMTP to $NewSMTP"
                    
                }
        
                # -- Email Address Removal --
                Write-Host "Checking Dynamic Distribution Group addresses on " $ThisDistroGroup.DisplayName
                foreach ($Address in $ThisDistroGroup.EmailAddresses) {
                    Write-Host "Checking " $Address
                    $AddressDomain = (($Address.Split(":")[1]).Split("@")[1]) # Extract domain from email address
        
                    if ($AddressDomain -eq $DomainToReplace) {
                        
                        if ($LoggingOnly -eq "Remove") {
                            $Command = "Set-DynamicDistributionGroup $GUID -EmailAddresses @{Remove='$Address'}"
                            iex $Command -WarningVariable Warning2 -ErrorVariable Error2
                            Start-Sleep -Seconds 1
        
                            # Append Warning and Error messages
                            if ($Warning2) { $WarningLog += "Email Removal Warning: " + ($Warning2 -join " | ") }
                            if ($Error2) { $ErrorLog += "Email Removal Error: " + ($Error2 -join " | ") }
                        }
        
                        $ActionValue = "Email address removed - $Address from $ThisDistroGroup.DisplayName"
                        Write-Host $ActionValue -ForegroundColor DarkRed
                        $UserObj | Add-Member -MemberType NoteProperty -Name "Action_EmailRemoval" -Value $ActionValue
                        
                    } else {
                        Write-Host "This address does not need to be removed"
                    }
                }
        
                # -- Append Exception Type & Message for each condition --
                if ($WarningLog -or $ErrorLog) {
                    $UserObj | Add-Member -MemberType NoteProperty -Name "ExceptionType" -Value "Warnings & Errors"
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Message" -Value (($WarningLog + $ErrorLog) -join " | ")
                }
            }
            catch {
                $ErrorLog += "General Dynamic Group Error: $($_.Exception.Message)"
                $UserObj | Add-Member -MemberType NoteProperty -Name "GeneralError" -Value $ErrorLog
            }
        }

        # Handling Group Mailboxes - Completed
        elseif (($RecipientType -eq "MailUniversalDistributionGroup") -and ($RecipientDetails -eq "GroupMailbox")) {
            try {
                # Fetch Group Mailbox (Unified Group) Information
                $ThisUnifiedGroup = Get-UnifiedGroup -Identity $GUID
                $UserObj | Add-Member -MemberType NoteProperty -Name "IsObjectSyncedAADC" -Value $ThisUnifiedGroup.IsDirSynced
                Write-Host "Checking Office 365 Group - " $ThisUnifiedGroup.DisplayName
        
                # Store errors and warnings
                $WarningLog = @()
                $ErrorLog = @()
        
                $ThisPrimarySMTP = $ThisUnifiedGroup.PrimarySMTPAddress
                $ThisPrimarySMTPDomain = $ThisPrimarySMTP.Split("@")[1]
                $ThisPrimarySMTPGroupNamePortion = $ThisPrimarySMTP.Split("@")[0]
                
                # -- Primary SMTP Update --
                if ($ThisPrimarySMTPDomain -eq $DomainToReplace) {
                    $NewSMTP = "$ThisPrimarySMTPGroupNamePortion@$DomainToUse"
                    
                    if ($LoggingOnly -eq "Remove") {
                        $Command = "Set-UnifiedGroup $GUID -PrimarySMTPAddress $NewSMTP"
                        iex $Command -WarningVariable Warning1 -ErrorVariable Error1
                        Start-Sleep -Seconds 1
        
                        # Append Warning and Error messages
                        if ($Warning1) { $WarningLog += "Primary SMTP Warning: " + ($Warning1 -join " | ") }
                        if ($Error1) { $ErrorLog += "Primary SMTP Error: " + ($Error1 -join " | ") }
                    }
        
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Action_SMTP" -Value "Primary SMTP changed from $ThisPrimarySMTP to $NewSMTP"
                    
                }
        
                # -- Email Address Removal --
                Write-Host "Checking Office 365 addresses on " $ThisUnifiedGroup.DisplayName
                foreach ($Address in $ThisUnifiedGroup.EmailAddresses) {
                    Write-Host "Checking " $Address
                    $AddressDomain = (($Address.Split(":")[1]).Split("@")[1]) # Extract domain from email address
        
                    if ($AddressDomain -eq $DomainToReplace) {
                        
                        if ($LoggingOnly -eq "Remove") {
                            $Command = "Set-UnifiedGroup $GUID -EmailAddresses @{Remove='$Address'}"
                            iex $Command -WarningVariable Warning2 -ErrorVariable Error2
                            Start-Sleep -Seconds 1
        
                            # Append Warning and Error messages
                            if ($Warning2) { $WarningLog += "Email Removal Warning: " + ($Warning2 -join " | ") }
                            if ($Error2) { $ErrorLog += "Email Removal Error: " + ($Error2 -join " | ") }
                        }
        
                        $ActionValue = "Email address removed - $Address from $ThisUnifiedGroup.DisplayName"
                        Write-Host $ActionValue -ForegroundColor DarkRed
                        $UserObj | Add-Member -MemberType NoteProperty -Name "Action_EmailRemoval" -Value $ActionValue
                        
                    } else {
                        Write-Host "This address does not need to be removed"
                    }
                }
        
                # -- Append Exception Type & Message for each condition --
                if ($WarningLog -or $ErrorLog) {
                    $UserObj | Add-Member -MemberType NoteProperty -Name "ExceptionType" -Value "Warnings & Errors"
                    $UserObj | Add-Member -MemberType NoteProperty -Name "Message" -Value (($WarningLog + $ErrorLog) -join " | ")
                }
        
            }
            catch {
                $ErrorLog += "General Office 365 Group Error: $($_.Exception.Message)"
                $UserObj | Add-Member -MemberType NoteProperty -Name "GeneralError" -Value $ErrorLog
            }
        }
        else {
            <# Action when all if and elseif conditions are false #>
			$ActionValue = "************Unrecognized recipient type!"
			Write-Host $ActionValue -foregroundcolor DarkRed
			$UserObj | Add-Member -Membertype NoteProperty -Name "Action" -Value $Actionvalue
        }
        $report += $UserObj       
    }

    Write-Host "All relevant email addresses removed" -ForegroundColor magenta

    #Let's log any remaining user accounts that are using the domain in some way, but don't have any Exchange attributes, and are therefore not Recipients
    #These cannot be updated as is - changing the UPN will still leave a proxy address in place containig the domain that must be removed
    #The account must either be deleted OR licensed with a mailbox, the address removed using EAC or Powershell, and then the license removed
    Write-Host "Let's check any user accounts that are using the domain but have no mail attributes" -foregroundcolor cyan
    $UsersWithDomain = Get-MsolUser -All -DomainName $DomaintoReplace
    ForEach ($User in $UsersWithDomain) {
	    #Is User not a recipient?
	    $UserCheck = Get-Recipient $User.UserPrincipalName -ErrorAction SilentlyContinue
	    If ($UserCheck -eq $NULL) {
		    Write-Host "User found with domain - " $User.DisplayName
		    $ThisUPN = $User.UserPrincipalName
		    $UserObj = New-Object PSObject
		    $UserObj | Add-Member -membertype Noteproperty -Name "Time" -value (get-date).ToString("yyyyMMdd-HH:mm:ss.fff")
		    $UserObj | Add-Member -membertype NoteProperty -Name "UPN" -Value $User.UserPrincipalName
		    $UserObj | Add-Member -membertype NoteProperty -Name "DisplayName" -Value $User.DisplayName
		    $UserObj | Add-Member -membertype NoteProperty -Name "RecipientType" -Value "User ONLY - Manual intervention required"
		    $UserObJ | Add-Member -membertype NoteProperty -Name "Action" -Value "***Manual intervention required for this object!"
		    $UserObj | Add-Member -Membertype NoteProperty -Name "IsObjectSyncedAADC" -Value $User.LastDirSyncTime
		    $report += $UserObj	
	    }
    }
    Write-Host "Finished checking all remaining users with the domain" -foregroundcolor Cyan

    # Export results
    $FileName = "C:\Users\adminuser\DomainRemovalReport-" + (Get-Date -Format "yyyyMMdd-HHmm") + ".csv"
    $report | Export-Csv -Path $FileName -NoTypeInformation


    Start-Sleep -Seconds 3
}

Write-Host "Script execution completed. Log file: $FileName" -ForegroundColor Green

# Disconnect session
Disconnect-ExchangeOnline -Confirm:$false
