Attribute VB_Name = "UDPSend"
Option Explicit

'=====================================
' === Simple Client Server Message ===
'=====================================
Sub UDPSend_KeepAlive()
  Owner.SeqNum1 = Owner.SeqNum1 + 1
  UDPPacket_SendRandom UDP_CMD_KEEP_ALIVE, True, Owner.SeqNum1
  DebugTxt "UDPSend", "CMD_KEEP_ALIVE"
End Sub

Sub UDPSend_ACK(ByVal intSeq As Integer)
  UDPPacket_SendRandom UDP_CMD_ACK, True, intSeq
  DebugTxt "UDPSend", "CMD_ACK for server reply" + Str$(intSeq)
End Sub

Sub UDPSend_ACKMessages()
  UDPPacket_SendRandom UDP_CMD_ACK_MESSAGES
  DebugTxt "UDPSend", "Send Got Message (CMD_ACK_MESSAGES)"
End Sub

Sub UDPSend_TextCode(ByVal txtCode As String)
  Dim Packet As String
  Owner.Command = UDP_CMD_SEND_TEXT_CODE
  Owner.SeqNum1 = Owner.SeqNum1 + 1
  Owner.Parameter = StrAppend(txtCode) + Dec_to_Hex(5, tInt)

  Packet = UDP_CreatePacketSeq(Owner, Owner.SeqNum1)
  UDPSock_Send Packet
  DebugTxt "UDPSend", "Send Text Code : " + txtCode
End Sub
'=============================================
' === Client to Client message thru server ===
'=============================================
Function UDPSend_OnlineMessage(MsgHead As MESSAGE_HEADER, Optional UINList As Variant, Optional NickList As Variant) As Integer
  Dim MsgCompiled As String, MaxContact As Integer, i As Integer
  With MsgHead
    Select Case .MSG_Type
      Case TYPE_MSG
        MsgCompiled = .MSG_Text
        DebugTxt "UDPSend", "Send Text Message to" + Str$(.lngUIN) + ". Message '" + .MSG_Text + "'"
        
      Case TYPE_URL
        MsgCompiled = .URL_Description + Chr$(&HFE) + .URL_Address
        DebugTxt "UDPSend", "Send URL Message to" + Str$(.lngUIN) + "." + vbCrLf + _
          "URL Address: " + .URL_Address + vbCrLf + _
          "Description: " + .URL_Description

      Case TYPE_AUTH_REQ
        MsgCompiled = _
          .AUTH_NickName + Chr$(&HFE) + _
          .AUTH_FirstName + Chr$(&HFE) + _
          .AUTH_LastName + Chr$(&HFE) + _
          .AUTH_Email + Chr$(&HFE) + _
          "1" + Chr$(&HFE) + .AUTH_Reason
        DebugTxt "UDPSend", "Request" + Str$(.lngUIN) + " for authorization, with reason '" + .AUTH_Reason + "'"
      
      Case TYPE_AUTH_DECLINE
        MsgCompiled = .AUTH_Reason
        DebugTxt "UDPSend", "Allow" + Str$(.lngUIN) + " to add us to his list"
        
      Case TYPE_AUTH_ACCEPT
        MsgCompiled = vbNullChar
        DebugTxt "UDPSend", "Allow" + Str$(.lngUIN) + " to add us to his list"
          
      Case TYPE_ADDED
        MsgCompiled = _
          .AUTH_NickName + Chr$(&HFE) + _
          .AUTH_FirstName + Chr$(&HFE) + _
          .AUTH_LastName + Chr$(&HFE) + _
          .AUTH_Email
        DebugTxt "UDPSend", "Notify" + Str$(.lngUIN) + " that we added him to our list"
        
      Case TYPE_CONTACT
        MaxContact = UBound(UINList) + 1
        MsgCompiled = Trim(Str$(MaxContact)) + Chr$(&HFE)
        For i = 0 To UBound(UINList)
          MsgCompiled = MsgCompiled + _
            Trim(Str$(UINList(i))) + Chr$(&HFE) + _
            NickList(i) + Chr(&HFE)
        Next i
        DebugTxt "UDPSend", "Sending" + Str$(MaxContact) + " contact list(s) to" + Str$(.lngUIN)
        
    End Select
    
    UDPSend_OnlineMessage = _
      UDPPacket_SendUIN(UDP_CMD_SEND_THRU_SRV, .lngUIN, "", _
      Dec_to_Hex(.MSG_Type, tInt) + StrAppend(MsgCompiled))
  End With
End Function

'=====================================================
'=== Login / Logout / Status / Registration Packet ===
'=====================================================
Function UDPSend_Login() As Integer
  Dim ParamTime As Long
      
  ParamTime = DateDiff("d", "1-1-1971", Now()) * 24 * 60 * 60
  ParamTime = ParamTime + Timer
  
  Owner.SessionID = CLng(Rnd(Timer) * &H3FFFFFFF)
  Owner.SeqNum1 = CInt(Rnd(Timer) * &H7FFF)
  Owner.SeqNum2 = 1
  Owner.Parameter = _
      Dec_to_Hex(ParamTime, tLong) + _
      Dec_to_Hex(ICQEngine.TCPListenPort, tLong) + _
      StrAppend(Owner.Password) + _
      "98000000" + _
      IP_to_Hex(ICQEngine.InternalIP) + _
      Dec_to_Hex(ICQEngine.ConnectionMethod, tByte) + _
      Dec_to_Hex(ICQEngine.InitialLoginStatus, tLong) + _
      "03000000" + _
      "00000000" + _
      "10009800"

  UDPSock_Connect
  UDPSend_Login = UDPPacket_SendString(UDP_CMD_LOGIN, Owner.Parameter)
  ICQEngine.ConnectionState = ICQ_Login
  DebugTxt "UDPSend", "Send Login Packet for UIN" + Str$(Owner.uin)
End Function

Function UDPSend_Login1() As Integer
  UDPSend_Login1 = UDPPacket_SendRandom(UDP_CMD_LOGIN_1)
  DebugTxt "UDPSend", "Send Login 1 Packet"
End Function

Sub UDPSend_Logout()
  UDPQueue_Reset
  UDPSend_TextCode "B_USER_DISCONNECTED"
  UDPSock_Close
  Notify_Disconnect
  ICQEngine.ConnectionState = ICQ_Login
  DebugTxt "UDPSend", "UDP Sock Disconnected" + vbCrLf
End Sub

Function UDPSend_RegNewUser() As Integer
  Owner.uin = 0
  Owner.SessionID = CLng(Rnd(Timer) * &H3FFFFFFF)
  Owner.SeqNum1 = Owner.SeqNum1 + 1
  Owner.SeqNum2 = 1
  
  UDPSend_RegNewUser = UDPPacket_SendString(UDP_CMD_REG_NEW_USER, StrAppend(Owner.Password), , _
    "A0000000612400000000A00000000000")
  ICQEngine.ConnectionState = ICQ_Register_New_User
  DebugTxt "UDPSend", "Registering new user using password: " + Owner.Password
End Function

Function UDPSend_ChangeStatus(OnlineStatus As LOGIN_ONLINE_STATUS) As Integer
  UDPSend_ChangeStatus = UDPPacket_SendString(UDP_CMD_STATUS_CHANGE, Dec_to_Hex(OnlineStatus, tLong))
  DebugTxt "UDPSend", "Change user status to " + Debug_OnlineStatusName(OnlineStatus)
End Function

'=================================================
'=== Contact / Visible / Invisible List Packet ===
'=================================================
Function UDPSend_AddToList(ByVal uin As Long) As Integer
  UDPSend_AddToList = UDPPacket_SendUIN(UDP_CMD_ADD_TO_LIST, uin)
  DebugTxt "UDP", "Adding user to contact list" + Str$(uin)
End Function

Function UDPSend_ContactList(Optional UINList) As Integer
  UDPSend_ContactList = UDPPacket_SendUINList(UDP_CMD_CONT_LIST, UINList)
  DebugTxt "UDPSend", "Send Contact List"
End Function

Function UDPSend_InvisList(Optional UINList) As Integer
  UDPSend_InvisList = UDPPacket_SendUINList(UDP_CMD_INVIS_LIST, UINList)
  DebugTxt "UDPSend", "Send Invisible List"
End Function

Function UDPSend_VisList(Optional UINList) As Integer
  UDPSend_VisList = UDPPacket_SendUINList(UDP_CMD_VIS_LIST, UINList)
  DebugTxt "UDPSend", "Send Visible List"
End Function

'==========================
'=== Search User Packet ===
'==========================
Function UDPSend_SearchReqUIN(ByVal uin As Long) As Integer
  UDPSend_SearchReqUIN = UDPPacket_SendUIN(UDP_CMD_SEARCH_UIN, uin)
  DebugTxt "UDPSend", "Searching user with UIN #" + Trim$(Str$(uin))
End Function

Function UDPSend_SearchReq(ByVal strNick As String, ByVal strFirst As String, ByVal strLast As String, ByVal strEmail As String) As Integer
  UDPSend_SearchReq = UDPPacket_SendString(UDP_CMD_SEARCH_USER, StrAppend(strNick, strFirst, strLast, strEmail))
  
  DebugTxt "UDPSend", "Searching for the following user" + vbCrLf + _
           "  Nick  : " + strNick + vbCrLf + _
           "  First : " + strFirst + vbCrLf + _
           "  Last  : " + strLast + vbCrLf + _
           "  Email : " + strEmail
End Function

'==========================
'=== User Detail Packet ===
'==========================

'** Request **
Function UDPSend_InfoReq(ByVal uin As Long) As Integer
  UDPSend_InfoReq = UDPPacket_SendUIN(UDP_CMD_INFO_REQ, uin)
  DebugTxt "UDPSend", "Requesting Info for UIN" + Str$(uin)
End Function

Function UDPSend_ExtInfoReq(ByVal uin As Long) As Integer
  UDPSend_ExtInfoReq = UDPPacket_SendUIN(UDP_CMD_EXT_INFO_REQ, uin)
  DebugTxt "UDPSend", "Requesting Extended Info for UIN" + Str$(uin)
End Function

Function UDPSend_MetaInfoReq(ByVal uin As Long) As Integer
  UDPSend_MetaInfoReq = UDPPacket_SendUIN(UDP_CMD_META_USER, uin, Dec_to_Hex(META_CMD_REQ_INFO, tInt))
  DebugTxt "UDPSend", "Requesting Meta User info for" + Str$(uin)
End Function

'** Update **
Function UDPSend_UpdateInfo(UserDetail As USER_DETAIL_INFO) As Integer
  With UserDetail
    UDPSend_UpdateInfo = UDPPacket_SendString(UDP_CMD_UPDATE_INFO, _
      StrAppend(.strNickname, .strFirstname, .strLastName, .strEmail))
    
    DebugTxt "UDPSend", "Updating Owner info to the following," + vbCrLf + _
             "  Nick  : " + .strNickname + vbCrLf + _
             "  First : " + .strFirstname + vbCrLf + _
             "  Last  : " + .strLastName + vbCrLf + _
             "  Email : " + .strEmail
  End With
End Function

Function UDPSend_UpdateAuthInfo(UserDetail As USER_DETAIL_INFO) As Integer
  UDPSend_UpdateAuthInfo = UDPPacket_SendString(UDP_CMD_UPDATE_AUTH, Dec_to_Hex(UserDetail.bAuthRequest, tLong))
  DebugTxt "UDPSend", "Updating user authorization info"
End Function

Function UDPSend_UpdateNewUserInfo(UserDetail As USER_DETAIL_INFO) As Integer
  With UserDetail
    UDPSend_UpdateNewUserInfo = UDPPacket_SendString( _
      UDP_CMD_NEW_USER_INFO, _
      StrAppend(.strNickname, .strFirstname, .strLastName, .strEmail), , "010101")
      
    DebugTxt "UDPSend", "Updating New User info"
  End With
End Function

Function UDPSend_UpdateMetaInfoMain(UserDetail As USER_DETAIL_INFO) As Integer
  
  With UserDetail
    UDPSend_UpdateMetaInfoMain = UDPPacket_SendString(UDP_CMD_META_USER, _
      StrAppend(.strNickname, .strFirstname, .strLastName, _
      .strEmail, .strEmail2, .strEmail3, .strCity, .strState, _
      .strPhone, .strFax, .strStreet, .strCellular), _
      Dec_to_Hex(META_CMD_SET_INFO, tInt), _
      Dec_to_Hex(.lngZip, tLong) + Dec_to_Hex(.intCountryCode, tInt) + _
      Dec_to_Hex(.intCountryStat, tInt) + Dec_to_Hex(.bEmailHide, tByte))
  
  DebugTxt "UDPSend", "Updating 'MAIN' Meta User info"
  End With
End Function

Function UDPSend_UpdateMetaInfoMore(UserDetail As USER_DETAIL_INFO) As Integer
  With UserDetail
  UDPSend_UpdateMetaInfoMore = UDPPacket_SendString(UDP_CMD_META_USER, _
    StrAppend(.strHomepageURL), _
    Dec_to_Hex(META_CMD_SET_MORE, tInt) + Dec_to_Hex(.byteAgeYear, tByte) + "0002", _
    Dec_to_Hex(.byteBirthYear, tByte) + Dec_to_Hex(.byteBirthMonth, tByte) + _
    Dec_to_Hex(.byteBirthDay, tByte) + "FFFFFF")
  End With
  
  DebugTxt "UDPSend", "Updating 'MORE' Meta User info"
End Function

Function UDPSend_UpdateMetaInfoAbout(UserDetail As USER_DETAIL_INFO) As Integer
  With UserDetail
    UDPSend_UpdateMetaInfoAbout = _
      UDPPacket_SendString(UDP_CMD_META_USER, _
      StrAppend(.strAboutInfo), _
      Dec_to_Hex(META_CMD_SET_ABOUT, tInt))
  
    DebugTxt "UDPSend", "Updating 'ABOUT' Meta User info"
  End With
End Function

Function UDPSend_UpdateMetaInfoSecurity(UserDetail As USER_DETAIL_INFO) As Integer
  With UserDetail
    UDPSend_UpdateMetaInfoSecurity = _
      UDPPacket_SendString(UDP_CMD_META_USER, _
      Dec_to_Hex(.bAuthRequest, tByte) + Dec_to_Hex(.bWebPresence, tByte) + Dec_to_Hex(.bPublishIP, tByte), _
      Dec_to_Hex(META_CMD_SET_SECURE, tInt))

    DebugTxt "UDPSend", "Updating User security info"
  End With
End Function

'========================================================================
'========================================================================
'======================== MAIN UDP PACKET ASSEMBLY ======================
'========================================================================
'========================================================================
Function UDPPacket_SendRandom(ByVal ClientCommand As UDP_CLIENT_COMMAND, Optional SpecialSeq As Boolean = False, Optional Seq As Integer) As Integer
  Dim Packet As String
  Owner.Command = ClientCommand
  Owner.Parameter = Dec_to_Hex(CLng(Rnd(Timer) * &H7FFFFFFF), tLong)

  If SpecialSeq = False Then
    Packet = UDP_CreatePacket(Owner)
    UDPQueue_Set Packet, Owner.SeqNum1
  Else
    Packet = UDP_CreatePacketSeq(Owner, Seq)
    If (ClientCommand <> UDP_CMD_ACK) And (ClientCommand <> UDP_CMD_SEND_TEXT_CODE) Then
      UDPQueue_Set Packet, Seq
    End If
  End If
  
  UDPSock_Send Packet
  UDPPacket_SendRandom = Owner.SeqNum1
End Function

Function UDPPacket_SendUIN(ByVal ClientCommand As UDP_CLIENT_COMMAND, ByVal uin As Long, Optional BeforeParameter As String = "", Optional AfterParameter As String = "") As Integer
  UDPPacket_SendUIN = UDPPacket_SendString(ClientCommand, Dec_to_Hex(uin, tLong), BeforeParameter, AfterParameter)
  UDPPacket_SendUIN = Owner.SeqNum1
End Function

Function UDPPacket_SendUINList(ByVal ClientCommand As UDP_CLIENT_COMMAND, Optional UINList)
  Dim Packet As String, AmountofUIN As Integer, i As Integer
  Owner.Command = ClientCommand
  
  If IsArray(UINList) Then
    AmountofUIN = UBound(UINList) - LBound(UINList) + 1
    Owner.Parameter = Dec_to_Hex(AmountofUIN, tByte)
    For i = LBound(UINList) To UBound(UINList)
      Owner.Parameter = Owner.Parameter + Dec_to_Hex(UINList(i), tLong)
    Next i
  Else
    Owner.Parameter = "00"
  End If

  Packet = UDP_CreatePacket(Owner)
  UDPSock_Send Packet
  UDPQueue_Set Packet, Owner.SeqNum1
  UDPPacket_SendUINList = Owner.SeqNum1
End Function

Function UDPPacket_SendString(ByVal ClientCommand As UDP_CLIENT_COMMAND, ByVal hexText As String, Optional BeforeParameter As String = "", Optional AfterParameter As String = "") As Integer
  Dim Packet As String
  Owner.Command = ClientCommand
  Owner.Parameter = BeforeParameter + hexText + AfterParameter

  Packet = UDP_CreatePacket(Owner)

  UDPSock_Send Packet
  UDPQueue_Set Packet, Owner.SeqNum1
  UDPPacket_SendString = Owner.SeqNum1
End Function
