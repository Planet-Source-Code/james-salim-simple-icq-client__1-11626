Attribute VB_Name = "modDebug"
Function DebugHexDump(HexPacket$)
    For i = 1 To Int(Len(HexPacket$) / 2) Step 16
        For j = 1 To 16
            TempOut$ = TempOut$ + PeekByte$(HexPacket$, i + j - 2) + " "
        Next j
        TempOut$ = TempOut$ + vbCrLf
    Next i
    DebugHexDump = TempOut$
End Function

Public Sub DebugTxt(strCategory As String, strText As String)
    If DEBUG_INFO Then
      ICQControl.rtfDebug.SelCharOffset = 0
      ICQControl.rtfDebug.SelText = "(" + Format(Date$, "dd mmm ") + Time$ + ")" + strCategory + " - " + strText + vbCrLf
    End If
End Sub

Public Sub DebugError(DbgError As ErrObject)
    DebugTxt "vbERROR", "The following error has occured:" + vbCrLf + _
             "  Error Number : " + Trim$(Str$(DbgError.Number)) + vbCrLf + _
             "  Description :  " + DbgError.Description
End Sub

Public Function Debug_SrvReplyName(SrvReplyCommand) As String
Select Case SrvReplyCommand
  Case UDP_SRV_ACK: SrvReplyName = "SRV_ACK"
  Case UDP_SRV_LOGIN_REPLY: SrvReplyName = "SRV_LOGIN_REPLY"
  Case UDP_SRV_USER_ONLINE: SrvReplyName = "SRV_USER_ONLINE"
  Case UDP_SRV_USER_OFFLINE: SrvReplyName = "SRV_USER_OFFLINE"
  Case UDP_SRV_USER_FOUND: SrvReplyName = "SRV_USER_FOUND"
  Case UDP_SRV_OFFLINE_MESSAGE: SrvReplyName = "SRV_OFFLINE_MESSAGE"
  Case UDP_SRV_END_OF_SEARCH: SrvReplyName = "SRV_END_OF_SEARCH"
  Case UDP_SRV_INFO_REPLY: SrvReplyName = "SRV_INFO_REPLY"
  Case UDP_SRV_EXT_INFO_REPLY: SrvReplyName = "SRV_EXT_INFO_REPLY"
  Case UDP_SRV_STATUS_UPDATE: SrvReplyName = "SRV_STATUS_UPDATE"
  Case UDP_SRV_X1: SrvReplyName = "SRV_X1"
  Case UDP_SRV_X2: SrvReplyName = "SRV_X2"
  Case UDP_SRV_UPDATE: SrvReplyName = "SRV_UPDATE"
  Case UDP_SRV_UPDATE_EXT: SrvReplyName = "SRV_UPDATE_EXT"
  Case UDP_SRV_NEW_UIN: SrvReplyName = "SRV_NEW_UIN"
  Case UDP_SRV_NEW_USER: SrvReplyName = "SRV_NEW_USER"
  Case UDP_SRV_QUERY: SrvReplyName = "SRV_QUERY"
  Case UDP_SRV_SYSTEM_MESSAGE: SrvReplyName = "SRV_SYSTEM_MESSAGE"
  Case UDP_SRV_ONLINE_MESSAGE: SrvReplyName = "SRV_ONLINE_MESSAGE"
  Case UDP_SRV_GO_AWAY: SrvReplyName = "SRV_GO_AWAY"
  Case UDP_SRV_TRY_AGAIN: SrvReplyName = "SRV_TRY_AGAIN"
  Case UDP_SRV_FORCE_DISCONNECT: SrvReplyName = "SRV_FORCE_DISCONNECT"
  Case UDP_SRV_MULTI_PACKET: SrvReplyName = "SRV_MULTI_PACKET"
  Case UDP_SRV_WRONG_PASSWORD: SrvReplyName = "SRV_WRONG_PASSWORD"
  Case UDP_SRV_INVALID_UIN: SrvReplyName = "SRV_INVALID_UIN"
  Case UDP_SRV_META_USER: SrvReplyName = "SRV_META_USER"
  Case UDP_SRV_RAND_USER: SrvReplyName = "SRV_RAND_USER"
  Case UDP_SRV_AUTH_UPDATE: SrvReplyName = "SRV_AUTH_UPDATE"
  Case Else: SrvReplyName = "Unknown" + Str$(SrvReplyCommand) + " " + Dec_to_Hex(SrvReplyCommand, tInt)
End Select
Debug_SrvReplyName = SrvReplyName
End Function

Function Debug_OnlineStatusName(OnlineStat As LOGIN_ONLINE_STATUS) As String
Select Case OnlineStat
  Case STATUS_ONLINE: Debug_OnlineStatusName = "ONLINE"
  Case STATUS_INVISIBLE: Debug_OnlineStatusName = "INVISIBLE"
  Case STATUS_NA: Debug_OnlineStatusName = "EXTENDED AWAY"
  Case STATUS_AWAY: Debug_OnlineStatusName = "AWAY"
  Case STATUS_DND: Debug_OnlineStatusName = "DO NOT DISTURB"
  Case STATUS_OCCUPIED: Debug_OnlineStatusName = "OCCUPIED"
  Case STATUS_CHAT: Debug_OnlineStatusName = "FREE FOR CHAT"
  Case STATUS_OFFLINE: Debug_OnlineStatusName = "OFFLINE"
End Select
End Function
