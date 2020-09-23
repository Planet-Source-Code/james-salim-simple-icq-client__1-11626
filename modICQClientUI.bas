Attribute VB_Name = "UI_Client"
Public lstContactList(10) As Long
Public lstVisibleList(5) As Long
Public lstInvisibleList(5) As Long

Sub UI_LoadLastSaved()
  TCPSock_Init
  With ICQEngine
    .UDPRemoteHost = "icq.mirabilis.com"
    .UDPRemotePort = 4000
    .InitialLoginStatus = STATUS_ONLINE
    .ConnectionMethod = LOGIN_NO_TCP
  End With
  
  Randomize Timer
  For i = 0 To 10
    lstContactList(i) = Int(Rnd(Timer) * &H6000000)
  Next i
End Sub

Sub UI_Init()
  With frmICQMain
    .Caption = Trim$(Str$(Owner.uin))
    For i = LBound(lstContactList) To UBound(lstContactList)
      .tvwContact.Nodes.Add , , , lstContactList(i), 1
    Next i
  End With
End Sub
