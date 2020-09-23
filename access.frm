VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0DAC1A01-B3C1-11D2-B00B-004033A11032}#3.0#0"; "TITULO.OCX"
Begin VB.Form acces 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Look Spy - ICQ 51964739"
   ClientHeight    =   2625
   ClientLeft      =   2505
   ClientTop       =   1830
   ClientWidth     =   6585
   Icon            =   "access.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin Titulo.anititulo anititulo1 
      Height          =   600
      Left            =   690
      TabIndex        =   6
      Top             =   615
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1058
   End
   Begin VB.Frame f 
      Height          =   585
      Index           =   1
      Left            =   210
      TabIndex        =   3
      Top             =   1965
      Width           =   6270
      Begin VB.TextBox txt_password 
         BackColor       =   &H80000001&
         Height          =   285
         Left            =   60
         TabIndex        =   4
         Text            =   "Status...."
         Top             =   195
         Width           =   6120
      End
   End
   Begin VB.Frame f 
      Height          =   735
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   1200
      Width           =   6240
      Begin VB.CommandButton cmd_end 
         Caption         =   "&Exit"
         Height          =   495
         Left            =   4920
         TabIndex        =   2
         Top             =   165
         Width           =   1215
      End
      Begin VB.CommandButton cmd_decript 
         Caption         =   "Choose .MDB"
         Height          =   495
         Left            =   105
         TabIndex        =   1
         Top             =   165
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog dialog 
         Left            =   1350
         Top             =   165
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Label l 
      Caption         =   "Look Spy - Access Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   105
      Width           =   5790
   End
End
Attribute VB_Name = "acces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_decript_Click()
dialog.Filter = "Access 97 (*.MDB)|*mdb"
dialog.ShowOpen
If Trim(access97(dialog.FileName)) <> "" Then 'Verify if this Database is password for access97
    txt_password = "Password is " & access97(dialog.FileName)
End If
If Len(Trim(txt_password)) = 11 Then txt_password = "Nothing Password....": MsgBox "Database is not password !": txt_password = "Nothing Password...."
End Sub


Private Sub cmd_end_Click()
MsgBox "Thank You !!"
End
End Sub


Private Sub Form_Load()
txt_password = "Please visity my Site for WEB: http://users.sti.com.br/psystem"
End Sub

Private Sub Timer1_Timer()

End Sub


