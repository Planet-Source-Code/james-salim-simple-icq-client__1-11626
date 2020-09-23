VERSION 5.00
Begin VB.Form frmUserPass 
   Caption         =   "Medievilz ICQ"
   ClientHeight    =   1230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3150
   Icon            =   "frmUserPass.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1230
   ScaleWidth      =   3150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   2160
      TabIndex        =   4
      Top             =   765
      Width           =   960
   End
   Begin VB.TextBox txtPass 
      Height          =   240
      Left            =   1170
      TabIndex        =   3
      Top             =   405
      Width           =   1950
   End
   Begin VB.TextBox txtUIN 
      Height          =   240
      Left            =   1170
      TabIndex        =   2
      Top             =   90
      Width           =   1950
   End
   Begin VB.Label Label2 
      Caption         =   "&Password"
      Height          =   240
      Left            =   90
      TabIndex        =   1
      Top             =   405
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "&UIN"
      Height          =   240
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   1275
   End
End
Attribute VB_Name = "frmUserPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  With Owner
    .uin = Val(txtUIN.Text)
    .Password = txtPass.Text
  End With
  
  Load frmICQMain
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Show
End Sub
