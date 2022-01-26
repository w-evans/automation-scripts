function get-identity {

    [CmdletBinding()]
    param(
    [Parameter(mandatory=$true)]
    [string]$racfid)

#### REQ MODULE CHECKING #################
    #import AD module
    Import-Module ActiveDirectory
    #AD loaded module check
    if (!(Get-Module ActiveDirectory)) { throw 'Active Directory Module not loaded..' }

######################################

    #list of domains we scan for user objects
    $domains = 'somedomain.com', 'somedomain.com', 'somedomain.com', 'somedomain.com', 'somedomain.com', 'somedomain.com'

    #ascii banner
    Invoke-RestMethod -uri "http://artii.herokuapp.com/make?text=YourCompany-IAM&font=thin" -DisableKeepAlive
    Write-Host 'Company Motto..' -ForegroundColor Red -BackgroundColor White


    for ($num = 0; $num -le $domains.length - 1; $num++) {
        try {
            
            #initial query
            $userObject = Get-ADuser -Identity $racfid -Properties * -Server $domains[$num]

            Write-Host "=====================" -ForegroundColor Gray -BackgroundColor Black
            Write-Host $domains[$num].toUpper() -ForegroundColor Black -BackgroundColor White
            Write-Host 'User Summary:' -ForegroundColor Yellow
            Write-Host 'displayName:' $userObject.displayName
            Write-Host 'Title:' $userObject.Title 
            Write-Host 'ENABLED:' $userObject.Enabled
            Write-Host 'LockedOut:' $userObject.LockedOut
            Write-Host 'PasswordExpired:' $userObject.PasswordExpired 
            Write-Host 'PasswordLastSet:' $userObject.PassWordLastSet  
            Write-Host 'PrimaryDomain:' $userObject.PrimaryADDomain
            Write-Host 'ObjectLocation:' $userObject.CanonicalName
            Write-Host 'mail:' $userObject.Mail 
            Write-Host 'mailNickName:' $userObject.mailNickName
            Write-Host 'UPN:' $userObject.userPrincipalName
            
            Write-Host ''
            Write-Host 'ProxyAddresses:' 
            Write-Output $userObject.proxyAddresses
            
            Write-Host ''
            Write-Host 'Groups:' 
            Write-Output $userObject.memberOf | Get-ADgroup -Server $domains[$num] | select -ExpandProperty Name | sort Name

            }
        catch { #nada
        }
    }
    Write-Host "======= DONE =========" -ForegroundColor Green -BackgroundColor White
}