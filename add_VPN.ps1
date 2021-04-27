$VPNaddress = 'pt-brwqkwngtbc.dynamic-m.com'
$PreSharedKey = 'D!gitalRail5'
$VPNname = 'PT-VPN'
$VPNtype = 'L2tp'


Add-VpnConnection -Name $VPNname -ServerAddress $VPNaddress -TunnelType $VPNtype -AllUserConnection -L2tpPsk $PreSharedKey -AuthenticationMethod Pap -Force


Get-VpnConnection
