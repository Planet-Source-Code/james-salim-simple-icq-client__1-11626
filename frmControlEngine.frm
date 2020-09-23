VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ICQControl 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "ICQ Debug Window"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   9120
   Icon            =   "frmControlEngine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   448
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   608
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtfDebug 
      Height          =   5685
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   10028
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmControlEngine.frx":0442
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   7605
      Top             =   2250
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEngine.frx":04FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEngine.frx":0658
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEngine.frx":07B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbarControl 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   6435
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6482
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6482
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrControl 
      Interval        =   1000
      Left            =   7470
      Top             =   1665
   End
   Begin MSWinsockLib.Winsock TCP 
      Left            =   7470
      Top             =   1170
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock UDP 
      Left            =   7470
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "ICQControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
  Me.Show
  rtfDebug.SelColor = vbWhite
  rtfDebug.SelText = "Visual ICQ Instant Messenger - James Salim" + vbCrLf + vbCrLf
End Sub

Private Sub Form_Resize()
  rtfDebug.Width = Me.ScaleWidth
  stbarControl.Top = 0
  rtfDebug.Height = stbarControl.Top
  rtfDebug.Height = stbarControl.Top
End Sub

Private Sub tmrControl_Timer()
  '******************************
  'Keep The builtin timer ticking
  '******************************
  icqTimer = icqTimer + 1
  If icqTimer = 32640 Then icqTimer = 0
    
  '**************************************
  'Make sure there is no message on queue
  '**************************************
  If ICQEngine.ConnectionState <> ICQ_Disconnected Then UDPQueue_CheckTime

  '*************************
  'KEEPALIVE Command Manager
  '*************************
  If ICQEngine.ConnectionState = ICQ_Connected Then
    KeepAliveTimer = KeepAliveTimer + 1
    If KeepAliveTimer > KeepAliveIntervalTime Then
      UDPSend_KeepAlive
      KeepAliveTimer = 0
    End If
  Else
    KeepAliveTimer = 0
  End If
  
  '*********************************************
  'Debug the time in the stbarControl status bar
  '*********************************************
  stbarControl.Panels(1).Text = "Timer :" + Str$(icqTimer Mod 60)
  stbarControl.Panels(2).Text = "KeepAlive Timer :" + Str$(KeepAliveTimer)
  Select Case ICQEngine.ConnectionState
    Case ICQ_Connected
      stbarControl.Panels(3).Picture = imgList.ListImages(1).ExtractIcon
      stbarControl.Panels(3).Text = "Connected"
    Case ICQ_Disconnected
      stbarControl.Panels(3).Picture = imgList.ListImages(2).ExtractIcon
      stbarControl.Panels(3).Text = "Disconnected"
    Case Else
      stbarControl.Panels(3).Picture = imgList.ListImages(3).ExtractIcon
      stbarControl.Panels(3).Text = "Login/Registering..."

  End Select
End Sub

Private Sub UDP_DataArrival(ByVal bytesTotal As Long)
  Dim SrvReply As UDP_SERVER_HEADER, _
      SubSrvReply As UDP_SERVER_HEADER, _
      strPacket As String, _
      TotalPacket As Integer, _
      ReadPos As Integer, i As Integer, PacketLength As Integer
      
    
  UDP.GetData strPacket
  strPacket = Str_to_Hex$(strPacket)
  SrvReply = UDPRecv_SplitServerHeader(strPacket)
        
  
  With SrvReply
    'Send acknowledgme for all server reply other than SRV_ACK
    If .Command <> UDP_SRV_ACK Then
      UDPSend_ACK .SeqNum1
    End If
      
    
    If .Command = UDP_SRV_MULTI_PACKET Then
      ReadPos = 1
      TotalPacket = Val("&H" + PeekByte(.Parameter, 0))
      DebugTxt "UDPRecv", "Multi Packet (" + Trim$(Str$(TotalPacket)) + " Packet)"
      
      For i = 1 To TotalPacket
        PacketLength = Val("&H" + Peek(.Parameter, ReadPos, 2))
        ReadPos = ReadPos + 2
        strPacket = hDump(Peek(.Parameter, ReadPos, PacketLength))
        ReadPos = ReadPos + PacketLength
        
        SubSrvReply = UDPRecv_SplitServerHeader(strPacket)
        UDPRecv_MainHandle SubSrvReply
      Next i
    Else
      UDPRecv_MainHandle SrvReply
    End If
  End With
End Sub



