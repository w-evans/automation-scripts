$VPNaddress = 'pt-brwqkwngtbc.dynamic-m.com'
$PreSharedKey = Read-Host 'enter pre-shared key' -MaskInput -AsSecureString
$VPNname = 'PT-VPN'
$VPNtype = 'L2tp'


Add-VpnConnection -Name $VPNname -ServerAddress $VPNaddress -TunnelType $VPNtype -AllUserConnection true -L2tpPs2 $PreSharedKey -AuthenticationMethod Pap -Force


