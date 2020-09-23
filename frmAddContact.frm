VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddContact 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add New Contact"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4905
   Icon            =   "frmAddContact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   1980
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   503
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8599
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add to List"
      Height          =   420
      Left            =   360
      TabIndex        =   2
      Top             =   1035
      Width           =   1365
   End
   Begin VB.TextBox txtUIN 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   585
      Width           =   4200
   End
   Begin VB.Label lblUIN 
      Caption         =   "ICQ User Identification Number:"
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   4290
   End
End
Attribute VB_Name = "frmAddContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
  If ICQEngine.ConnectionState = ICQ_Connected Then
    uin = Val(txtUIN.Text)
    frmICQMain.tvwContact.Nodes.Add , tvwLast, , uin, 1
    UDPSend_AddToList uin
    stBar.Panels(1).Text = txtUIN.Text + " were added to the contact list"
    txtUIN.Text = ""
  Else
    MsgBox "You need to be at least connected to use this function"
  End If
End Sub

Private Sub Form_Load()
  Me.Show
End Sub

Private Sub txtUIN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    cmdAdd_Click
  End If
End Sub
