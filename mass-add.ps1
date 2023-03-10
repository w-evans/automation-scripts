function Mass-Add {
    #any amount of users to any amount of groups in any amount of domains
    $exceldata = Import-CSV "C:\Users.csv"
    
    $users = Import-CSV 'C:\Users.csv' | Select -ExpandProperty userID
    $domains = Import-CSV 'C:\Users.csv' | Select -ExpandProperty domain
    $groups = Import-CSV 'C:\Users.csv' | Select -ExpandProperty group
    
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
