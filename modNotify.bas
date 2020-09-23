Attribute VB_Name = "modNotify"
Public Sub Notify_UDPAck(SequenceNumber As Integer)
End Sub

Sub Notify_NewUIN(uin As Long)
  UDPSock_Connect
  Owner.uin = .uin
  UDPSend_Login
End Sub

Sub Notify_LoggedIn(Optional ErrorMessage As String = "")
  If ErrorMessage = "" Then
    UDPSend_ContactList lstContactList()
    UDPSend_VisList
    UDPSend_InvisList
    frmICQMain.StatusBar.Panels(1).Text = Debug_OnlineStatusName(ICQEngine.InitialLoginStatus)
  Else
    MsgBox "The following error has occured during login:" + vbCrLf + vbCrLf + ErrorMessage, vbExclamation + vbOKOnly, "Login Error"
  End If
End Sub

Sub Notify_ContactStatusChange(ContactDetail As CONTACT_DETAIL, Optional SetTCPInfo As Boolean = False)
  With ContactDetail
    For i = 0 To frmICQMain.tvwContact.Nodes.Count - 1
      If Val(frmICQMain.tvwContact.Nodes(i + 1).Text) = .lngUIN Then
        If .lngStatus = STATUS_OFFLINE Then
          frmICQMain.tvwContact.Nodes(i + 1).Image = 1
        Else
          frmICQMain.tvwContact.Nodes(i + 1).Image = 2
        End If
      End If
    Next i
  End With
End Sub

Sub Notify_RecvMessage(MSGRecv As MESSAGE_HEADER)
  Load frmMessageRecv
  frmMessageRecv.stBar.Panels(1).Text = "Last message received at " + MSGRecv.MSG_Date + " " + MSGRecv.MSG_Time
  
  With frmMessageRecv.rtfMessage
  Select Case MSGRecv.MSG_Type
    Case TYPE_MSG
      .SelCharOffset = 0
      .SelColor = &H808000
      .SelItalic = True
      .SelText = Trim$(Str$(MSGRecv.lngUIN)) + " says:" + vbCrLf
      .SelItalic = False
      
      .SelColor = vbBlack
      .SelBold = True
      .SelText = MSGRecv.MSG_Text
      .SelText = vbCrLf + vbCrLf
      .SelBold = False
    Case TYPE_URL
      .SelCharOffset = 0
      .SelColor = &H808000
      .SelItalic = True
      .SelText = Trim$(Str$(MSGRecv.lngUIN)) + " send URL address:" + vbCrLf
      .SelItalic = False
      
      .SelColor = vbBlack
      .SelBold = True
      .SelText = "URL Address: " + MSGRecv.URL_Address
      .SelText = vbCrLf
      .SelText = "Description: " + MSGRecv.URL_Description
      .SelText = vbCrLf + vbCrLf
      .SelBold = False
    Case TYPE_ADDED
      .SelCharOffset = 0
      .SelColor = &H808000
      .SelItalic = True
      .SelText = Trim$(Str$(MSGRecv.lngUIN)) + " added you to his/her contact list"
      .SelText = vbCrLf + vbCrLf
      .SelItalic = False
    Case Else
      .SelCharOffset = 0
      .SelColor = &H808000
      .SelItalic = True
      .SelText = Trim$(Str$(MSGRecv.lngUIN)) + " sends you a different type of message that is not currently handled in this version of Medievilz ICQ Client."
      .SelText = vbCrLf + vbCrLf
      .SelItalic = False

    End Select
  End With
End Sub

Sub Notify_Disconnect()
    For i = LBound(lstContactList) To UBound(lstContactList)
      frmICQMain.tvwContact.Nodes(i + 1).Image = 1
    Next i
    frmICQMain.StatusBar.Panels(1).Text = Debug_OnlineStatusName(STATUS_OFFLINE)

End Sub

