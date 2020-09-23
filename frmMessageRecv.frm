VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMessageRecv 
   Caption         =   "Message Received"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   Icon            =   "frmMessageRecv.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   3435
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8308
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfMessage 
      Height          =   2715
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   4789
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMessageRecv.frx":030A
   End
End
Attribute VB_Name = "frmMessageRecv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Me.Show
End Sub

Private Sub Form_Resize()
  stBar.Top = 0
  rtfMessage.Width = Me.ScaleWidth
  rtfMessage.Height = stBar.Top
End Sub
