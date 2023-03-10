function Mass-Add {
    #any amount of users to any amount of groups in any amount of domains
    $exceldata = Import-CSV "C:\Users.csv"
    
    $users = $exceldata | Select -ExpandProperty userID
    $domains = $exceldata | Select -ExpandProperty domain
    $groups = $exceldata | Select -ExpandProperty group
    
    ForEach ($user in $users) {  
        ForEach ($domain in $domains) {
            ForEach ($group in $groups) { 
                Try {
                    Add-ADGroupMember -Server $domain -Id $group -Members $user.toString()
                    Write-Host $user 'added to' $group 'in' $domain
                }
                Catch {
                    Write-Host 'error:' $domain ':' $user ':' $group ':' $_
                }
            }
        }      
    }
}
