Attribute VB_Name = "modMain"
Sub Main()
  Load ICQControl
  Load frmUserPass
End Sub

Public Sub UDPSock_Send(ByVal Packet As String)
    ICQControl.UDP.SendData Packet
End Sub

Public Sub UDPSock_Connect(Optional LocalUDPPort As Long = 0)
    ICQControl.UDP.RemoteHost = ICQEngine.UDPRemoteHost
    ICQControl.UDP.RemotePort = ICQEngine.UDPRemotePort
    ICQControl.UDP.LocalPort = LocalUDPPort
    ICQControl.UDP.Connect
    ICQEngine.UDPLocalPort = ICQControl.UDP.LocalPort
    DebugTxt "UDPSock", "Connecting to " + ICQEngine.UDPRemoteHost + ":" + Trim$(Str$(ICQEngine.UDPRemotePort)) + " ..."
End Sub

Public Sub UDPSock_Close()
    ICQControl.UDP.Close
    DebugTxt "UDPSock", "Connection Closed"
End Sub

Public Sub TCPSock_Init()
  'ICQControl.TCP.Listen
  ICQEngine.InternalIP = ICQControl.TCP.LocalIP
  ICQEngine.TCPListenPort = ICQControl.TCP.LocalPort
End Sub
