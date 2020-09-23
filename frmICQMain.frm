VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmICQMain 
   Caption         =   "Medievilz ICQ"
   ClientHeight    =   4500
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3120
   Icon            =   "frmICQMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   208
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgList 
      Left            =   2385
      Top             =   3510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICQMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICQMain.frx":0896
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Text            =   "OFFLINE"
            TextSave        =   "OFFLINE"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwContact 
      Height          =   3840
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   6773
      _Version        =   393217
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "&Save As..."
         Enabled         =   0   'False
      End
      Begin VB.Menu sep00 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogin 
         Caption         =   "&Login"
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "My &Status"
         Begin VB.Menu mnuStatusOnline 
            Caption         =   "&Online"
         End
         Begin VB.Menu mnuStatusInvisible 
            Caption         =   "&Invisible"
         End
         Begin VB.Menu mnuStatusAway 
            Caption         =   "&Away"
         End
         Begin VB.Menu mnuStatusNA 
            Caption         =   "&Extended Away"
         End
         Begin VB.Menu mnuStatusOccupied 
            Caption         =   "&Busy"
         End
         Begin VB.Menu mnuStatusDND 
            Caption         =   "&Do Not Disturb"
         End
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenRecvFolder 
         Caption         =   "Open &Received Files folder"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "&Properties"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep3a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuSend 
      Caption         =   "&Contacts"
      Begin VB.Menu muListManager 
         Caption         =   "&Lists Manager"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "&Search User"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuViewProfile 
         Caption         =   "View &Profile"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAddContact 
         Caption         =   "&Add User"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSendMessage 
         Caption         =   "Send &Message"
      End
      Begin VB.Menu mnuSendURL 
         Caption         =   "Send &Web URL"
      End
      Begin VB.Menu mnuSendFile 
         Caption         =   "Send &File"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSendChat 
         Caption         =   "Request &Chat"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSendEmail 
         Caption         =   "Send E-&Mail"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCheckMail 
         Caption         =   "Check E-mail"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmICQMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NewUser As Boolean
Public ImageAnimateNumber As Integer

Private Sub Form_Load()
  Me.Show
  UI_LoadLastSaved
  UI_Init
End Sub

Private Sub Form_Resize()
  StatusBar.Top = 0
  tvwContact.Width = frmICQMain.ScaleWidth
  If StatusBar.Top > 0 Then
    tvwContact.Height = StatusBar.Top
  Else
    tvwContact.Height = 0
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  UDPSend_Logout
End Sub

Private Sub mnuAddContact_Click()
  Load frmAddContact
End Sub

Private Sub mnuClose_Click()
  If ICQEngine.ConnectionState = ICQ_Connected Then UDPSend_Logout
  End
End Sub

Private Sub mnuLogin_Click()
  If mnuLogin.Caption = "&Login" Then
    mnuLogin.Caption = "&Logout"
    TCPSock_Init
    UDPSend_Login
  Else
    mnuLogin.Caption = "&Login"
    UDPSend_Logout
  End If
End Sub

Private Sub mnuSendMessage_Click()
  Load frmMessageSend
End Sub

Private Sub mnuSendURL_Click()
  Load frmMessageSend
End Sub

Private Sub mnuStatusAway_Click()
  UDPSend_ChangeStatus STATUS_AWAY
  frmICQMain.StatusBar.Panels(1).Text = Debug_OnlineStatusName(STATUS_AWAY)
End Sub

Private Sub mnuStatusDND_Click()
  UDPSend_ChangeStatus STATUS_DND
  frmICQMain.StatusBar.Panels(1).Text = Debug_OnlineStatusName(STATUS_DND)
End Sub

Private Sub mnuStatusInvisible_Click()
  UDPSend_ChangeStatus STATUS_INVISIBLE
  frmICQMain.StatusBar.Panels(1).Text = Debug_OnlineStatusName(STATUS_INVISIBLE)
End Sub

Private Sub mnuStatusNA_Click()
  UDPSend_ChangeStatus STATUS_NA
  frmICQMain.StatusBar.Panels(1).Text = Debug_OnlineStatusName(STATUS_NA)
End Sub

Private Sub mnuStatusOccupied_Click()
  UDPSend_ChangeStatus STATUS_OCCUPIED
  frmICQMain.StatusBar.Panels(1).Text = Debug_OnlineStatusName(STATUS_OCCUPIED)
End Sub

Private Sub mnuStatusOnline_Click()
  UDPSend_ChangeStatus STATUS_ONLINE
  frmICQMain.StatusBar.Panels(1).Text = Debug_OnlineStatusName(STATUS_ONLINE)
End Sub
