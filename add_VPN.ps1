$VPNaddress = 'pt-brwqkwngtbc.dynamic-m.com'
$PreSharedKey = 'D!gitalRail5'
$VPNname = 'PT-VPN'
$VPNtype = 'L2tp'


Add-VpnConnection -Name $VPNname -ServerAddress $VPNaddress -TunnelType $VPNtype -AllUserConnection true -L2tpPs2 $PreSharedKey -AuthenticationMethod Pap -Force


