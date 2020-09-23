Attribute VB_Name = "UDPReceive"
'========================================================================
'======================== MAIN UDP PACKET HANDLE ========================
'========================================================================



Sub UDPRecv_MainHandle(SrvReply As UDP_SERVER_HEADER)
  Dim UserDetail As USER_DETAIL_INFO, lngUIN As Long, MSGRecv As MESSAGE_HEADER
  
  With SrvReply
    If .Version <> 5 Then
      DebugTxt "UDPRecv", "ERROR - Invalid Server Packet version" + Str$(.Version) + " received. Attempting to handle ..."
    End If
    
    Select Case .Command
      Case UDP_SRV_ACK:
        DebugTxt "UDPRecv", "Server acknowledged packet" + Str$(.SeqNum1)
        UDPQueue_ACK .SeqNum1
      Case UDP_SRV_X1
        DebugTxt "UDPRecv", "Acknowleged UDP_SRV_X1"
      Case UDP_SRV_X2
        DebugTxt "UDPRecv", "Acknowleged UDP_SRV_X2"
        UDPSend_ACKMessages
        
      Case UDP_SRV_NEW_UIN
        DebugTxt "UDPRecv", "New User Registration Accepted. The UIN is" + Str$(.uin)
        Notify_NewUIN SrvReply.uin
        UDPSock_Close
      Case UDP_SRV_LOGIN_REPLY:
        ICQEngine.ExternalIP = Hex_to_IP(Peek(.Parameter, 0, tLong))
        ICQEngine.ConnectionState = ICQ_Connected
        DebugTxt "UDPRecv", "Login Successful, UIN" + Str$(.uin) + " IP" + ICQEngine.ExternalIP
        UDPSend_Login1
        Notify_LoggedIn

      Case UDP_SRV_WRONG_PASSWORD
        DebugTxt "UDPRecv", "Wrong Password"
        Notify_LoggedIn "Unable to connect to the ICQ network." + vbCrLf + _
                        "Possible reason is wrong password or uin."
                        
      Case UDP_SRV_INVALID_UIN
        DebugTxt "UDPRecv", "Invalid UIN"
        Notify_LoggedIn "Unable to connect to the ICQ network." + vbCrLf + _
                        "Possible reason is wrong password or uin."
                        
      Case UDP_SRV_TRY_AGAIN
        UDPSock_Close
        UDPSock_Connect
        UDPSend_Login
        DebugTxt "UDPRecv", "SRV_TRY_AGAIN Server is busy, trying to login again..."
        
      Case UDP_SRV_GO_AWAY
        UDPSend_Logout
        DebugTxt "UDPRecv", "SRV_GO_AWAY, disconnected from server"
        
      Case UDP_SRV_INFO_REPLY:      UDPRecv_HandleInfoReply .Parameter, .Command
      Case UDP_SRV_EXT_INFO_REPLY:  UDPRecv_HandleInfoReply .Parameter, .Command
      Case UDP_SRV_META_USER:       UDPRecv_HandleInfoReply .Parameter, .Command
        
      Case UDP_SRV_USER_ONLINE:     UDPRecv_HandleStatusChange .Parameter, .Command
      Case UDP_SRV_USER_OFFLINE:    UDPRecv_HandleStatusChange .Parameter, .Command
      Case UDP_SRV_STATUS_UPDATE:   UDPRecv_HandleStatusChange .Parameter, .Command

      Case UDP_SRV_USER_FOUND:      UDPRecv_HandleSearch .Parameter, .Command
      Case UDP_SRV_END_OF_SEARCH:   UDPRecv_HandleSearch .Parameter, .Command
        
      Case UDP_SRV_OFFLINE_MESSAGE: UDPRecv_HandleMessageReply .Parameter, .Command
      Case UDP_SRV_ONLINE_MESSAGE:  UDPRecv_HandleMessageReply .Parameter, .Command
              
      Case Else:    DebugTxt "UDPRecv", "ERROR - Unrecognized Server Reply. Command " + Debug_SrvReplyName(.Command)
    End Select
  End With
End Sub


Sub UDPRecv_HandleSearch(ByVal Parameter As String, ByVal Command As Integer)

End Sub

Function UDPRecv_SplitServerHeader(ByVal strPacket As String) As UDP_SERVER_HEADER
  Dim SrvReply As UDP_SERVER_HEADER
  
  With SrvReply
    .Version = Val("&H" + Peek(strPacket, 0, 2))
    Select Case .Version
      Case 5
        .SessionID = Val("&H" + Peek(strPacket, 3, 4))
        .SeqNum1 = Val("&H" + Peek(strPacket, 9, tInt))
        .SeqNum2 = Val("&H" + Peek(strPacket, 11, tInt))
        .uin = Val("&H" + Peek(strPacket, 13, tLong))
        .Command = Val("&H" + Peek(strPacket, 7, tInt))
        .Parameter = CutTextL(strPacket, 21 * 2)
      Case 4
        .Command = Val("&H" + Peek(strPacket, 6, tInt))
        .SeqNum1 = Val("&H" + Peek(strPacket, 8, tInt))
        .SeqNum2 = Val("&H" + Peek(strPacket, 10, tInt))
        .uin = Val("&H" + Peek(strPacket, 12, tLong))
        .Parameter = CutTextL(strPacket, 20 * 2)
      Case 3
        .Command = Val("&H" + Peek(strPacket, 2, tInt))
        .SeqNum1 = Val("&H" + Peek(strPacket, 4, tInt))
        .SeqNum2 = Val("&H" + Peek(strPacket, 6, tInt))
        .uin = Val("&H" + Peek(strPacket, 8, tLong))
        .Parameter = CutTextL(strPacket, 16 * 2)
      Case 2
        .Command = Val("&H" + Peek(strPacket, 2, tInt))
        .SeqNum1 = Val("&H" + Peek(strPacket, 4, tInt))
        .Parameter = CutTextL(strPacket, 6 * 2)
    End Select
  End With

  UDPRecv_SplitServerHeader = SrvReply
End Function

Sub UDPRecv_HandleInfoReply(ByVal Parameter As String, ByVal Command As UDP_SERVER_REPLY)
  Select Case Command
    Case UDP_SRV_INFO_REPLY
        DebugTxt "UDPRecv", "Info Request reply received"
    Case UDP_SRV_EXT_INFO_REPLY
      DebugTxt "UDPRecv", "Extended info Request reply received"
    Case UDP_SRV_META_USER
      DebugTxt "UDPRecv", "Meta User info Request reply received"
  End Select
End Sub

Sub UDPRecv_HandleMessageReply(ByVal Parameter As String, ByVal Command As UDP_SERVER_REPLY)
  Dim MSGRecv As MESSAGE_HEADER, TempSplit
  
  With MSGRecv
    Select Case Command
      Case UDP_SRV_OFFLINE_MESSAGE
        DebugTxt "UDPRecv", "SRV_OFFLINE_MESSAGE received"
        .MSG_Date = Trim(Str$(Val("&H" + Peek(Parameter, 7, tByte)))) + "-" + _
                     Trim(Str$(Val("&H" + Peek(Parameter, 6, tByte)))) + "-" + _
                     Trim(Str$(Val("&H" + Peek(Parameter, 4, tInt))))
        .MSG_Time = Trim(Str$(Val("&H" + Peek(Parameter, 8, tByte)))) + ":" + _
                    Trim(Str$(Val("&H" + Peek(Parameter, 9, tByte))))
        .MSG_Type = Val("&H" + Peek(Parameter, 10, tInt))
        .MSG_Text = Hex_to_Str(CutTextL(Parameter, 28))
      Case UDP_SRV_ONLINE_MESSAGE
        DebugTxt "UDPRecv", "SRV_ONLINE_MESSAGE received"
        .MSG_Date = Format(Date$, "dd-mm-yyyy")
        .MSG_Time = Format(Time$, "hh:mm")
        .MSG_Type = Val("&H" + Peek(Parameter, 4, tInt))
        .MSG_Text = Hex_to_Str(CutTextL(Parameter, 16))
    End Select
    .lngUIN = Val("&H" + Peek(Parameter, 0, tLong))
    
    Select Case .MSG_Type
      Case TYPE_URL
        TempSplit = StrSplitbyChar(.MSG_Text, Chr$(&HFE), 1)
        .URL_Description = TempSplit(0)
        .URL_Address = TempSplit(1)
      Case TYPE_AUTH_REQ
        TempSplit = StrSplitbyChar(.MSG_Text, Chr$(&HFE), 5)
        .AUTH_NickName = TempSplit(0)
        .AUTH_FirstName = TempSplit(1)
        .AUTH_LastName = TempSplit(2)
        .AUTH_Email = TempSplit(3)
        .AUTH_Reason = TempSplit(5)
      Case TYPE_AUTH_DECLINE
        TempSplit = StrSplitbyChar(.MSG_Text, Chr$(&HFE), 0)
        .AUTH_Reason = TempSplit(0)
      Case TYPE_ADDED
        TempSplit = StrSplitbyChar(.MSG_Text, Chr$(&HFE), 3)
        .AUTH_NickName = TempSplit(0)
        .AUTH_FirstName = TempSplit(1)
        .AUTH_LastName = TempSplit(2)
        .AUTH_Email = TempSplit(3)
      Case TYPE_WEBPAGER
        TempSplit = StrSplitbyChar(.MSG_Text, Chr$(&HFE), 5)
        .AUTH_NickName = TempSplit(0)
        .AUTH_FirstName = TempSplit(1)
        .AUTH_LastName = TempSplit(2)
        .AUTH_Email = TempSplit(3)
        .MSG_Text = TempSplit(5)
      Case TYPE_EXPRESS
        TempSplit = StrSplitbyChar(.MSG_Text, Chr$(&HFE), 5)
        .AUTH_NickName = TempSplit(0)
        .AUTH_FirstName = TempSplit(1)
        .AUTH_LastName = TempSplit(2)
        .AUTH_Email = TempSplit(3)
        .MSG_Text = TempSplit(5)
    End Select
  End With
  Notify_RecvMessage MSGRecv
End Sub

Sub UDPRecv_HandleStatusChange(ByVal Parameter As String, ByVal Command As UDP_SERVER_REPLY)
  Dim ContactDetail As CONTACT_DETAIL
  With ContactDetail
    Select Case Command
      Case UDP_SRV_USER_ONLINE
          .lngUIN = CLng(Val("&H" + Peek(Parameter, 0, tLong)))
          .TCP_ExternalIP = Hex_to_IP(Peek(Parameter, 4, tLong))
          .TCP_ExternalPort = Val("&H" + Peek(Parameter, 8, tLong))
          .TCP_InternalIP = Hex_to_IP(Peek(Parameter, 12, tLong))
          .TCP_FLAG = Val("&H" + PeekByte(Parameter, 16))
          .lngStatus = Val("&H" + Peek(Parameter, 17, tInt))
          .TCP_Version = Val("&H" + Peek(Parameter, 21, tLong))
          Notify_ContactStatusChange ContactDetail, True
          DebugTxt "UDPRecv", "USER_ONLINE" + Str$(.lngUIN) + " set status to " + Debug_OnlineStatusName(.lngStatus)
      Case UDP_SRV_USER_OFFLINE
          .lngUIN = Val("&H" + Peek(Parameter, 0, tLong))
          .lngStatus = STATUS_OFFLINE
          Notify_ContactStatusChange ContactDetail
          DebugTxt "UDPRecv", "USER_OFFLINE" + Str$(.lngUIN)
      Case UDP_SRV_STATUS_UPDATE
          .lngUIN = Val("&H" + Peek(Parameter, 0, tLong))
          .lngStatus = Val("&H" + Peek(Parameter, 4, tInt))
          Notify_ContactStatusChange ContactDetail, True
          DebugTxt "UDPRecv", "USER_STATUS" + Str$(.lngUIN) + " change status to " + Debug_OnlineStatusName(.lngStatus)
    End Select
  End With
End Sub
