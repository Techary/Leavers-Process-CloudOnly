#####################################
## Author: James Tarran // Techary ##
#####################################
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

        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser

        import-module ExchangeOnlineManagement

        }

    Function Get-MSonline{

        Set-PSRepository -Name "PSgallery" -InstallationPolicy Trusted

        Install-Module MSOnline -Scope CurrentUser

        }

    Function Get-AzureAD{

        Set-PSRepository -Name "PSgallery" -InstallationPolicy Trusted

        Install-Module AzureAD -Scope CurrentUser

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

    if (Get-MsolUser -UserPrincipalName $global:upn -ErrorAction SilentlyContinue) 
        {

            write-host -NoNewline "Searching for $global:upn"
            CountDown 3

            Write-host "User found!"
            start-sleep 1

        }

    else 
        {

            write-host -NoNewline "Searching for $global:upn"
            CountDown 3

            write-host "User not found, try again" 
            get-upn

        }

    }

# Removes licences and converts the string ID to something we're more familiar with. 
function removeLicences {

    $AssignedLicences = (get-MsolUser -UserPrincipalName $global:upn).licenses.AccountSkuId

    Invoke-WebRequest -uri https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv -outfile .\licences.csv | Out-Null
    $licences = import-csv .\licences.csv
    remove-item .\licences.csv -Force

    [System.Collections.ArrayList]$script:UFLicences = @()

    foreach ($Assignedlicence in $Assignedlicences)
        {

            $Assignedlicence = $Assignedlicence.Split(':')[-1]

            foreach ($licence in $licences)
                {

                    if ($Assignedlicence -like $licence.String_id)
                        {

                            if($script:UFLicences -notcontains $licence.Product_Display_name)
                                {

                                    $script:UFLicences = $script:UFLicences += $licence.Product_Display_name
                                    
                                }

                        }

                }

        }

    (get-MsolUser -UserPrincipalName $global:upn).licenses.AccountSkuId |
    foreach
        {

            Set-MsolUserLicense -UserPrincipalName $global:upn -RemoveLicenses $_

        }

}

# Removes UPN from global address list. 
function Remove-GAL {

        Do 
            { 
                cls

                print-TecharyLogo
                
                Write-host "**********************"
                Write-host "** Remove from GAL  **"
                Write-Host "**********************"
                
                    $script:hideFromGAL  = Read-Host "Do you want to remove the mailbox from the global address list? ( y / n ) "
                    Switch ($script:hideFromGAL)
                    {
                        Y 
                            { 
                                Set-Mailbox -Identity $global:upn -HiddenFromAddressListsEnabled $true

                                Write-host "$global:upn has been hidden"

                                remove-distributionGroups

                            }

                        N 
                            { 

                                remove-distributionGroups
                            
                            }

                        Default { "You didn't enter an expect response, you idiot." }
                    }
            }
        until ($script:hideFromGAL  -eq 'y' -or $script:hideFromGAL  -eq 'n')

}

# Lists all distri's and prompts to remove UPN from them or not
function remove-distributionGroups {

    cls

    print-TecharyLogo

    Write-host "*************************"
    Write-host "** Distribution groups **"
    Write-host "*************************"

    $mailbox = Get-Mailbox -Identity $global:upn
    $DN=$mailbox.DistinguishedName
    $Filter = "Members -like ""$DN"""
    $DistributionGroupsList = Get-DistributionGroup -ResultSize Unlimited -Filter $Filter
    Write-host `n
    Write-host "Listing all Distribution Groups:"
    Write-host `n
    $DistributionGroupsList | ft

    Do 
        {

            $script:removeDisitri = Read-Host "Do you want to remove $global:upn from all distribution groups ( y / n )?"
            Switch ($script:removeDisitri)
            {
                Y 
                    {  
                        ForEach ($item in $DistributionGroupsList) 
                            {

                                Remove-DistributionGroupMember -Identity $item.PrimarySmtpAddress -Member $global:upn -BypassSecurityGroupManagerCheck -Confirm:$false
                                Write-host "Successfully removed"
                    
                            }
                        Add-Autoreply

                    }
                Default { "You didn't enter an expect response, you idiot." }

                N { Add-Autoreply }
            }

        }
    until ($script:removeDisitri -eq 'y' -or $script:removeDisitri -eq 'n')

}

# Prompts to add an auto response or not
function Add-Autoreply {
    Do 
    { 
        cls

        print-TecharyLogo
        
        Write-Host "***************"
        Write-host "** Autoreply **"
        Write-host "***************"
        
        $script:autoreply = Read-Host "Do you want to add an auto-reply to $global:upn's mailbox? ( y / n / dog ) " 
        Switch ($script:autoreply) 
        { 
            Y { $oof = Read-Host "Enter auto-reply"

                Set-MailboxAutoReplyConfiguration -Identity $global:upn -AutoReplyState Enabled -ExternalMessage "$oof" -InternalMessage "$oof"
                write-host "Auto-reply added."
                Add-MailboxPermissions 
              } 
            N { Add-MailboxPermissions } 
            Default { "You didn't enter an expect response, you idiot." }
            Dog {   
                    write-host "  __      _"
                    write-host  "o'')}____//"
                    write-host  " `_/      )"
                    write-host  " (_(_/-(_/"
                    start-sleep 5
                    Add-Autoreply

                }

        }

        }
    until ($script:autoreply -eq 'y' -or $script:autoreply -eq 'n' -or $script:autoreply -eq 'Dog')
}

# Prompts to add mailbox permissions or not
function Add-MailboxPermissions{
    Do 
        { 
            cls

            print-TecharyLogo
            
            Write-host "*************************"
            Write-host "** Mailbox Permissions **"
            Write-Host "*************************"
        
            $script:mailboxpermissions = Read-Host "Do you want anyone to have access to this mailbox? ( y / n ) "
            Switch ($script:mailboxpermissions)
                {
                    Y { $WhichUser = Read-Host "Enter the E-mail address of the user that should have access to this mailbox "

                        add-mailboxpermission -identity $global:upn -user $WhichUser -AccessRights FullAccess

                        Write-host "Malibox permisions for $whichUser have been added"

                        write-result

                        }

                    N {write-result}

                    Default { "You didn't enter an expect response, you idiot." }
                }
        }
    until ($script:mailboxpermissions -eq 'y' -or $script:mailboxpermissions -eq 'n')
    
}

function write-result {

        write-host "You have done the following:"

        write-host "`nRemoved the following licences:"
        $script:UFLicences

        if ($script:hideFromGAL -eq 'N')
            {
                write-host -ForegroundColor Yellow "`nYou have not hidden $global:upn from the global address list."
            }
        else
            {
                write-host -ForegroundColor Green  "`nYou have hidden $global:upn from the global address list."
            }

        if($script:removeDisitri -eq 'N')
            {
                write-host -ForegroundColor Yellow "`nYou have not removed $global:upn from all distribution groups"
            }
        else
            {
                write-host -ForegroundColor Green "`nYou have removed $global:upn from any distribution groups."
            }

        if ($script:autoreply -eq 'N')
            {
                write-host -ForegroundColor Yellow "`nYou have not added an autoreply to $global:upn"
            }
        else 
            {
                write-host -ForegroundColor Green "`nYou have added an autoreply to $global:upn"
            }

        if($script:mailboxpermissions -eq 'N')
            {
                write-host -ForegroundColor Yellow "`nYou have not added any mailbox permissions to $global:upn"
            }
        else
            {
                write-host -ForegroundColor Green "`nYou have added mailbox permissions to $global:upn"
            }


        pause


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

connect-365

get-upn

removeLicences

Set-Mailbox $global:upn -Type Shared

Remove-GAL