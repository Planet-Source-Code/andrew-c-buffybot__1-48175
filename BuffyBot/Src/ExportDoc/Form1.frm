VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Exports Generator"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   540
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2070
      Top             =   1125
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&GO!"
      Height          =   495
      Left            =   3165
      TabIndex        =   4
      Top             =   2055
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   495
      Left            =   3165
      TabIndex        =   3
      Top             =   1305
      Width           =   1215
   End
   Begin VB.TextBox SRCDir 
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   915
      Width           =   4125
   End
   Begin VB.Label Label2 
      Caption         =   "SRC Dir:"
      Height          =   240
      Left            =   255
      TabIndex        =   1
      Top             =   645
      Width           =   3450
   End
   Begin VB.Label Label1 
      Caption         =   "This program will scan the buffybot source code and create a file documenting the exported scripting functions. Press GO:"
      Height          =   885
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   4470
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CommonDialog1.Filter = "VB Project | *.vbp"
CommonDialog1.ShowOpen
SRCDir.Text = CommonDialog1.Filename
End Sub

Private Sub Command2_Click()
ProcessVBP CommonDialog1.Filename
End Sub
