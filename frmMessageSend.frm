VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMessageSend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sending Message"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmMessageSend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmURLMSG 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2220
      Left            =   225
      TabIndex        =   6
      Top             =   1395
      Width           =   4245
      Begin VB.TextBox txtURLDesc 
         Height          =   285
         Left            =   45
         TabIndex        =   10
         Top             =   1215
         Width           =   4110
      End
      Begin VB.TextBox txtURLAdd 
         Height          =   285
         Left            =   90
         TabIndex        =   8
         Top             =   405
         Width           =   4110
      End
      Begin VB.Label Label2 
         Caption         =   "URL Description:"
         Height          =   240
         Left            =   90
         TabIndex        =   9
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "URL Address:"
         Height          =   240
         Left            =   90
         TabIndex        =   7
         Top             =   135
         Width           =   1500
      End
   End
   Begin VB.Frame frmTextMSG 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2220
      Left            =   225
      TabIndex        =   4
      Top             =   1395
      Width           =   4290
      Begin VB.TextBox txtMessage 
         Height          =   2175
         Left            =   0
         MaxLength       =   450
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   0
         Width           =   4245
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   3465
      TabIndex        =   3
      Top             =   3780
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   2715
      Left            =   135
      TabIndex        =   2
      Top             =   990
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   4789
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Text Message"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "URL Message"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmUIN 
      Caption         =   "Target UIN"
      Height          =   735
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   4470
      Begin VB.TextBox txtUIN 
         Height          =   285
         Left            =   135
         TabIndex        =   1
         Top             =   315
         Width           =   4200
      End
   End
End
Attribute VB_Name = "frmMessageSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
  Dim MSGSend As MESSAGE_HEADER
  If ICQEngine.ConnectionState = ICQ_Connected Then
    MSGSend.lngUIN = Val(txtUIN.Text)
    Select Case TabStrip.SelectedItem
      Case "Text Message"
        MSGSend.MSG_Type = TYPE_MSG
        MSGSend.MSG_Text = txtMessage.Text
      Case "URL Message"
        MSGSend.MSG_Type = TYPE_URL
        MSGSend.URL_Address = txtURLAdd.Text
        MSGSend.URL_Description = txtURLDesc.Text
    End Select
    UDPSend_OnlineMessage MSGSend
  Else
    MsgBox "You need to be at least connected to use this function"
  End If
End Sub

Private Sub Form_Load()
  Me.Show
  frmURLMSG.Visible = False
End Sub

Private Sub TabStrip_Click()
  If TabStrip.SelectedItem = "Text Message" Then
    frmURLMSG.Visible = False
   'frmTextMSG.Visible = True
  Else
    frmURLMSG.Visible = True
    'frmTextMSG.Visible = False
  End If
End Sub
