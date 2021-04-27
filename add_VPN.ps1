$VPNaddress = 'pt-brwqkwngtbc.dynamic-m.com'
<<<<<<< HEAD
$PreSharedKey = Read-Host 'enter pre-shared key' -MaskInput -AsSecureString
=======
$PreSharedKey = ''
>>>>>>> caf1fe67dfe77d3b53656f07f6682b813ffb5f6a
$VPNname = 'PT-VPN'
$VPNtype = 'L2tp'


Add-VpnConnection -Name $VPNname -ServerAddress $VPNaddress -TunnelType $VPNtype -AllUserConnection -L2tpPsk $PreSharedKey -AuthenticationMethod Pap -Force


Get-VpnConnection
