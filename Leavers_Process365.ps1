#####################################
## Author: James Tarran // Techary ##
#####################################

# ---------------------- ELEVATE ADMIN ---------------------- 

param([switch]$Elevated)

function Test-Admin {
  $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
  $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if ((Test-Admin) -eq $false)  {
    if ($elevated) 
    {
        # tried to elevate, did not work, aborting
    } 
    else {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
}

exit

}


# ---------------------- USED FUNCTIONS ----------------------

# Prints 'Techary' in ASCII
function print-TecharyLogo {
        
    $logo = "
     _______        _                      
    |__   __|      | |                     
       | | ___  ___| |__   __ _ _ __ _   _ 
       | |/ _ \/ __| '_ \ / _`` | '__| | | |
       | |  __/ (__| | | | (_| | |  | |_| |
       |_|\___|\___|_| |_|\__,_|_|   \__, |
                                      __/ |
                                     |___/ 
"

write-host -ForegroundColor Green $logo

}

# Checks to see if AzureAD, MSOnline, and Exchangeonelinemanagement are installed. 
# If not, installs them. Then connects to 365 online
function connect-365 {

    function invoke-mfaConnection{

        Connect-ExchangeOnline

        Import-Module AzureAD

        import-module MSOnline

        import-module ExchangeOnlineManagement

        Connect-MsolService

        }

    function Get-ExchangeOnlineManagement{

        Set-PSRepository -Name "PSgallery" -InstallationPolicy Trusted

        Install-Module -Name ExchangeOnlineManagement

        import-module ExchangeOnlineManagement

        }

    Function Get-MSonline{

        Set-PSRepository -Name "PSgallery" -InstallationPolicy Trusted

        Install-Module MSOnline

        }

    Function Get-AzureAD{

        Set-PSRepository -Name "PSgallery" -InstallationPolicy Trusted

        Install-Module AzureAD

            }

    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        write-host " "
        write-host "Exchange online Management exists"
    } 
    else {
        Write-host "Exchange Online Management module does not exist. Please ensure powershell is running as admin. Attempting to download..."
        Get-ExchangeOnlineManagement
    }


    if (Get-Module -ListAvailable -Name MSOnline) {
        write-host "MSOnline exists"
    } 
    else {
        Write-host "MSOnline module does not exist. Please ensure powershell is running as admin. Attempting to download..."
        Get-MSOnline
    }


    if (Get-Module -ListAvailable -Name AzureAD) {
        write-host "AzureAD exists"
    } 
    else {
        Write-host "AzureAD module does not exist. Please ensure powershell is running as admin. Attempting to download..."
        Get-AzureAD
    }

invoke-mfaConnection

}
# If the UPN is not specified as a parameter, asks for it here.
# Then checks if the user exists in 365. 
function get-upn {

    $global:upn = read-host "Input UPN"

    if (Get-MsolUser -UserPrincipalName $global:upn -ErrorAction SilentlyContinue) {Write-host "User found..."
     $global:upn
    }

    else {write-host "User not found, try again" 
        get-upn
}

    }

# Removes licences and converts the string ID to something we're more familiar with.
# Then writes a warning that this licence will need to be removed from the 365 portal. 
function removeLicences {

    $AssignedLicences = (get-MsolUser -UserPrincipalName $global:upn).licenses.AccountSkuId

    Invoke-WebRequest -uri https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv -outfile .\licences.csv | Out-Null
    $licences = import-csv .\licences.csv
    remove-item .\licences.csv -Force

    [System.Collections.ArrayList]$UFLicences = @()

    foreach ($Assignedlicence in $Assignedlicences)
        {

            $Assignedlicence = $Assignedlicence.Split(':')[-1]

            foreach ($licence in $licences)
                {

                    if ($Assignedlicence -like $licence."String_ id")
                        {

                            if($UFLicences -notcontains $licence.Product_Display_name)
                                {

                                    $UFLicences = $UFLicences += $licence.Product_Display_name
                                    
                                }

                        }

                }

        }

    Write-Warning "The following licence(s) will be removed from the user. Please remove from the 365 tenant:"
    $UFLicences

    pause

    (get-MsolUser -UserPrincipalName $global:upn).licenses.AccountSkuId |
    foreach{
        Set-MsolUserLicense -UserPrincipalName $global:upn -RemoveLicenses $_
    }

}

# Removes UPN from global address list. 
function Remove-GAL {

        Do { cls

        print-TecharyLogo
        
        Write-host "**********************"
        Write-host "** Remove from GAL  **"
        Write-Host "**********************"
        
            $MailboxPermission = Read-Host "Do you want to remove the mailbox from the global address list? ( y / n ) "
            Switch ($mailboxPermission)
            {
                Y { Set-Mailbox -Identity $upn -HiddenFromAddressListsEnabled $true

                    Write-host "$upn has been hidden"

                    remove-distributionGroups

                   }

                N { 
                    remove-distributionGroups
                    
                   }

                Default { "You didn't enter an expect response, you idiot." }
            }
        }
        until ($AutoReply -eq 'y' -or $AutoReply -eq 'n')

}

# Lists all distri's and prompts to remove UPN from them or not
function remove-distributionGroups{

    cls

    print-TecharyLogo

    Write-host "*************************"
    Write-host "** Distribution groups **"
    Write-host "*************************"

    $mailbox = Get-Mailbox -Identity $upn
    $DN=$mailbox.DistinguishedName
    $Filter = "Members -like ""$DN"""
    $DistributionGroupsList = Get-DistributionGroup -ResultSize Unlimited -Filter $Filter
    Write-host `n
    Write-host "Listing all Distribution Groups:"
    Write-host `n
    $DistributionGroupsList | ft

    Do {
        $RemoveDistri = Read-Host "Do you want to remove $upn from all distribution groups ( y / n )?"
        Switch ($RemoveDistri)
        {
            Y {  ForEach ($item in $DistributionGroupsList) {
                    Remove-DistributionGroupMember -Identity $item.PrimarySmtpAddress –Member $upn –BypassSecurityGroupManagerCheck -Confirm:$false
                    Write-host "Successfully removed"
                
                                            }
                Add-Autoreply
                                }
            Default { "You didn't enter an expect response, you idiot." }

            N { Add-Autoreply }
            }
        }
            until ($RemoveDistri -eq 'y' -or $RemoveDistri -eq 'n') 
}

# Prompts to add an auto response or not
function Add-Autoreply {
    Do { cls

        print-TecharyLogo
        
        Write-Host "***************"
        Write-host "** Autoreply **"
        Write-host "***************"
        
        $AutoReply = Read-Host "Do you want to add an auto-reply to $upn's mailbox? ( y / n / dog ) " 
        Switch ($AutoReply) 
        { 
            Y { $oof = Read-Host "Enter auto-reply"

        Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Enabled -ExternalMessage "$oof" -InternalMessage "$oof"
        write-host "Auto-reply added."
        Add-MailboxPermissions 
              } 
            N { Add-MailboxPermissions } 
            Default { "You didn't enter an expect response, you idiot." }
            Dog {   write-host "   __      _"
                    write-host  "o'')}____//"
                    write-host  " `_/      )"
                    write-host  " (_(_/-(_/"
                    start-sleep 5
                    Add-Autoreply
                    }
            }
        }
        until ($AutoReply -eq 'y' -or $AutoReply -eq 'n' -or $AutoReply -eq 'Dog')
}

# Prompts to add mailbox permissions or not
function Add-MailboxPermissions{
    Do { cls

        print-TecharyLogo
        
        Write-host "*************************"
        Write-host "** Mailbox Permissions **"
        Write-Host "*************************"
        
            $MailboxPermission = Read-Host "Do you want anyone to have access to this mailbox? ( y / n ) "
            Switch ($mailboxPermission)
            {
                Y { $WhichUser = Read-Host "Enter the E-mail address of the user that should have access to this mailbox "

                    add-mailboxpermission -identity $upn -user $WhichUser -AccessRights FullAccess

                    Write-host "Malibox permisions for $whichUser  have been added"

                    Disconnect-ExchangeOnline -Confirm:$false

                    exit

                    }

                N { Write-host "Ending Session..." 

                    Disconnect-ExchangeOnline -Confirm:$false

                    exit

                    }

                Default { "You didn't enter an expect response, you idiot." }
            }
        }
        until ($AutoReply -eq 'y' -or $AutoReply -eq 'n')
}

# Adds a feature to call '...' when waiting
function CountDown() {
    param($timeSpan)

    while ($timeSpan -gt 0)
  {
    Write-Host '.' -NoNewline
    $timeSpan = $timeSpan - 1
    Start-Sleep -Seconds 1
  }
}

# ---------------------- START SCRIPT ----------------------

$global:upn = $null

print-TecharyLogo

Write-host "Updating modules. This may take some time..."

countDown -timeSpan 25

Update-Module

connect-365

$upn = get-upn

removeLicences

Set-Mailbox $upn -Type Shared

Remove-GAL