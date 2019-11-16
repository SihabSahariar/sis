VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   5130
   ClientLeft      =   9165
   ClientTop       =   4230
   ClientWidth     =   8865
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "<<<<<<<<<<<<<<<<<<BACK>>>>>>>>>>>>>>>>>>"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   4560
      Width           =   8535
   End
   Begin VB.Image Image1 
      Height          =   4395
      Left            =   2040
      Picture         =   "Form2.frx":0000
      Top             =   0
      Width           =   4800
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

