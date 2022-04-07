function Mass-Add {
    #any amount of users to any amount of groups in any amount of domains
    $users = Import-CSV 'C:\Users.csv' | Select -ExpandProperty userID
    $domains = 'domain1', 'domain2', 'domain3', 'domain4', 'domain5'
    $groups = 'group1', 'group2', 'group3', 'group4', 'group5'
    
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
