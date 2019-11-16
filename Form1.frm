VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   Caption         =   $"Form1.frx":0000
   ClientHeight    =   9450
   ClientLeft      =   -240
   ClientTop       =   -135
   ClientWidth     =   12735
   FillStyle       =   2  'Horizontal Line
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":00F3
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":6945
   ScaleHeight     =   9450
   ScaleWidth      =   12735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H003C3C3C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   3720
      ScaleHeight     =   4095
      ScaleWidth      =   8535
      TabIndex        =   29
      Top             =   2400
      Visible         =   0   'False
      Width           =   8535
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H007BA329&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         ScaleHeight     =   615
         ScaleWidth      =   8535
         TabIndex        =   30
         Top             =   0
         Width           =   8535
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Software Development"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   2880
            TabIndex        =   31
            Top             =   120
            Width           =   4335
         End
      End
      Begin prj_si.Button Button12 
         Height          =   735
         Left            =   3000
         TabIndex        =   32
         Top             =   3240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1296
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "OK"
         Enabled         =   -1  'True
         Color           =   -2147483630
         Picture         =   "Form1.frx":19AF3F
         Picture         =   "Form1.frx":1A6871
         Select_Image    =   "Form1.frx":1B21A3
         FixedSingle     =   0
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   585
         Left            =   3960
         MouseIcon       =   "Form1.frx":1BDAD5
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":1BDC27
         Stretch         =   -1  'True
         ToolTipText     =   "Contact at Facebook"
         Top             =   2400
         Width           =   570
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Programmer and Designer : Sihab Sahariar Sizan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   480
         TabIndex        =   33
         Top             =   1800
         Width           =   7575
      End
   End
   Begin prj_si.Button Button10 
      Height          =   375
      Left            =   11880
      TabIndex        =   24
      Top             =   135
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      Enabled         =   -1  'True
      Color           =   -2147483630
      Picture         =   "Form1.frx":1C1735
      Picture         =   "Form1.frx":1C1B0C
      Select_Image    =   "Form1.frx":1C1EE3
      FixedSingle     =   0
   End
   Begin prj_si.Button Button6 
      Height          =   975
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Students Info"
      Enabled         =   -1  'True
      Color           =   -2147483630
      Picture         =   "Form1.frx":1C22D3
      Picture         =   "Form1.frx":1CDC05
      Select_Image    =   "Form1.frx":1D9537
      FixedSingle     =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   4320
      Picture         =   "Form1.frx":1E4E69
      ScaleHeight     =   5655
      ScaleWidth      =   8415
      TabIndex        =   4
      Top             =   3840
      Width           =   8415
      Begin VB.TextBox Text1 
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   2280
         TabIndex        =   10
         Top             =   4200
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000006&
         DataField       =   "Name"
         DataSource      =   "Data1"
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   1200
         TabIndex        =   9
         Top             =   0
         Width           =   4215
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000006&
         DataField       =   "Phone"
         DataSource      =   "Data1"
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000006&
         DataField       =   "Class"
         DataSource      =   "Data1"
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   1200
         TabIndex        =   7
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000006&
         DataField       =   "Roll"
         DataSource      =   "Data1"
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   1200
         TabIndex        =   6
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Data Data1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Connect         =   "Access 2000;"
         DatabaseName    =   "F:\Clu Nebula\db1.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001E1E1E&
         Height          =   825
         Left            =   960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Table1"
         Top             =   2520
         Width           =   4575
      End
      Begin prj_si.Button Button1 
         Height          =   615
         Left            =   6120
         TabIndex        =   5
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Add New"
         Enabled         =   -1  'True
         Color           =   -2147483630
         Picture         =   "Form1.frx":270F7B
         Picture         =   "Form1.frx":27C8AD
         FixedSingle     =   0
      End
      Begin prj_si.Button Button2 
         Height          =   615
         Left            =   6120
         TabIndex        =   11
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Clear"
         Enabled         =   -1  'True
         Color           =   -2147483630
         Picture         =   "Form1.frx":2881DF
         Picture         =   "Form1.frx":293B11
         FixedSingle     =   0
      End
      Begin prj_si.Button Button3 
         Height          =   615
         Left            =   6120
         TabIndex        =   12
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Delete"
         Enabled         =   -1  'True
         Color           =   -2147483630
         Picture         =   "Form1.frx":29F443
         Picture         =   "Form1.frx":2AAD75
         FixedSingle     =   0
      End
      Begin prj_si.Button Button4 
         Height          =   615
         Left            =   6120
         TabIndex        =   18
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Save"
         Enabled         =   -1  'True
         Color           =   -2147483630
         Picture         =   "Form1.frx":2B66A7
         Picture         =   "Form1.frx":2C1FD9
         FixedSingle     =   0
      End
      Begin prj_si.Button Button5 
         Height          =   615
         Left            =   6120
         TabIndex        =   19
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Go"
         Enabled         =   -1  'True
         Color           =   -2147483630
         Picture         =   "Form1.frx":2CD90B
         Picture         =   "Form1.frx":2D923D
         FixedSingle     =   0
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name :-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001E1E1E&
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone :-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   0
         TabIndex        =   16
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Class :-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001E1E1E&
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Roll :-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001E1E1E&
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Search  :-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001E1E1E&
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   4200
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   5040
      Picture         =   "Form1.frx":2E4B6F
      ScaleHeight     =   2775
      ScaleWidth      =   6735
      TabIndex        =   3
      Top             =   720
      Width           =   6735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5160
      TabIndex        =   0
      Top             =   10800
      Width           =   13095
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   5400
         Top             =   720
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   8520
         TabIndex        =   2
         Text            =   "Text7"
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Text            =   "Text6"
         Top             =   480
         Width           =   3975
      End
   End
   Begin prj_si.Button Button7 
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   5640
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Parents Info"
      Enabled         =   -1  'True
      Color           =   -2147483630
      Picture         =   "Form1.frx":326161
      Picture         =   "Form1.frx":331A93
      Select_Image    =   "Form1.frx":33D3C5
      FixedSingle     =   0
   End
   Begin prj_si.Button Button8 
      Height          =   975
      Left            =   120
      TabIndex        =   22
      Top             =   6840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Result Maker"
      Enabled         =   -1  'True
      Color           =   -2147483630
      Picture         =   "Form1.frx":348CF7
      Picture         =   "Form1.frx":354629
      Select_Image    =   "Form1.frx":35FF5B
      FixedSingle     =   0
   End
   Begin prj_si.Button Button9 
      Height          =   975
      Left            =   120
      TabIndex        =   23
      Top             =   8040
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Exit"
      Enabled         =   -1  'True
      Color           =   -2147483630
      Picture         =   "Form1.frx":36B88D
      Picture         =   "Form1.frx":3771BF
      Select_Image    =   "Form1.frx":382AF1
      FixedSingle     =   0
   End
   Begin prj_si.Button Button11 
      Height          =   375
      Left            =   11280
      TabIndex        =   28
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "?"
      Enabled         =   -1  'True
      Color           =   -2147483630
      Picture         =   "Form1.frx":38E423
      Picture         =   "Form1.frx":399D55
      Select_Image    =   "Form1.frx":3A5687
      FixedSingle     =   0
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "VS 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001E1E1E&
      Height          =   375
      Left            =   3600
      TabIndex        =   27
      Top             =   9120
      Width           =   615
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007BA329&
      Height          =   375
      Left            =   2520
      TabIndex        =   26
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007BA329&
      Height          =   375
      Left            =   1440
      TabIndex        =   25
      Top             =   3050
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const MF_BYPOSITION = &H400&
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&
Private Const HWND_TOPMOST = -1

Option Explicit
Dim bTrans As Byte ' The level of transparency (0 - 255)
Dim lOldStyle As Long
''''''''''''''''''''''
Dim XX As Integer
Dim YY As Integer

Private Sub Button10_Click()
End
End Sub

Private Sub Button11_Click()
Picture3.Visible = True
End Sub

Private Sub Button12_Click()
Picture3.Visible = False
End Sub

Private Sub Button6_Click()
MsgBox "For full version contact programmer.", vbInformation, "OK"
End Sub

Private Sub Button8_Click()
MsgBox "For full version contact programmer.", vbInformation, "OK"
End Sub

Private Sub Form_Load()
    bTrans = 0
   lOldStyle = SetWindowLong(Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED)
    Timer1.Enabled = True
    Timer1.Interval = 25
    Dim hMenu As Long
    hMenu = GetSystemMenu(Me.hwnd, False)
    DeleteMenu hMenu, 6, MF_BYPOSITION
    'SetWindowPos Me.hwnd, HWND_TOPMOST, 10, 10, 500, 500, 0
    Me.Show
    Data1.DatabaseName = App.Path & "\db1.mdb"
    Label7.Caption = Time
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    XX = x
    YY = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
    Me.Left = Me.Left - XX + x
    Me.Top = Me.Top - YY + Y
    End If
End Sub

Private Sub Image1_Click()
Shell "explorer.exe " & "http://facebook.com/sizan.first/"
End Sub

Private Sub Timer1_Timer()
bTrans = bTrans + 5
If bTrans >= 255 Then
    Timer1.Enabled = False
    Exit Sub
End If
SetLayeredWindowAttributes Me.hwnd, 0, bTrans, LWA_ALPHA
End Sub
Private Sub Button1_Click()
MsgBox "Entry a new record.", vbInformation, "Entry"
Data1.Recordset.AddNew
Text2.SetFocus
End Sub

Private Sub Button2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub Button3_Click()
On Error GoTo error
Data1.Recordset.Delete
Data1.Refresh
MsgBox "Record deleted successfully.", vbInformation, "OK"
error:
Data1.Refresh
MsgBox "No record found!", vbInformation, "OK"
End Sub

Private Sub Button4_Click()
Data1.Recordset.Update
MsgBox "Record saved successfully.", vbInformation, "OK"
End Sub

Private Sub Button5_Click()
Dim content
    content = Trim(Text1.Text) & "*"
    content = "Name like '" & content & "'"
    If Text1.Text <> "" Then
        Data1.Recordset.FindFirst content
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()
Data1.Refresh
End Sub



Private Sub Command7_Click()
End
End Sub

Private Sub Command8_Click()

End Sub

Private Sub Button7_Click()
MsgBox "For full version contact programmer.", vbInformation, "OK" 'Picture2.Picture = App.Path & "UI\b2.bmp"
End Sub

Private Sub Button9_Click()
End

End Sub



