$VPNaddress = 'pt-brwqkwngtbc.dynamic-m.com'
$PreSharedKey = Read-Host "enter Pre-Shared Key: "
$VPNname = 'PT-VPN'
$VPNtype = 'L2tp'

#Remove ALL existing VPNs
Get-VpnConnection -AllUserConnection | Remove-VpnConnection -Force

#Add our VPN connection, will prompt user for PSK
Add-VpnConnection -Name $VPNname -ServerAddress $VPNaddress -TunnelType $VPNtype -AllUserConnection -L2tpPsk $PreSharedKey -AuthenticationMethod Pap -Force

#Verify that our VPN has been added
Get-VpnConnection -AllUserConnection
