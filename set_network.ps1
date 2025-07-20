# Set static IP address for Ethernet adapter
New-NetIPAddress -InterfaceAlias "Ethernet" -IPAddress "192.168.15.101" -PrefixLength 24 -DefaultGateway "192.168.15.254"
