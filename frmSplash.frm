VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3210
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   6585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   2640
      Top             =   1800
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   2880
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   11
      Scrolling       =   1
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   6120
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   0
      Picture         =   "frmSplash.frx":08D6
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   465
   End
   Begin VB.Image imgLogo 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1665
      Left            =   480
      Picture         =   "frmSplash.frx":11A0
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      Height          =   1935
      Index           =   1
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Picture"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2520
      TabIndex        =   7
      Top             =   240
      Width           =   2010
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4320
      TabIndex        =   6
      Top             =   1680
      Width           =   1965
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright© 2001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   4320
      TabIndex        =   5
      Top             =   2280
      Width           =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   6120
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   360
      Y2              =   1560
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   6120
      X2              =   2520
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   6120
      X2              =   6120
      Y1              =   360
      Y2              =   1560
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00004080&
      BorderWidth     =   2
      X1              =   2400
      X2              =   2400
      Y1              =   240
      Y2              =   1680
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C0E0FF&
      BorderWidth     =   2
      X1              =   6240
      X2              =   6240
      Y1              =   240
      Y2              =   1680
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00004080&
      BorderWidth     =   2
      X1              =   2400
      X2              =   6240
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00C0E0FF&
      BorderWidth     =   2
      X1              =   2400
      X2              =   6240
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VIEWER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   810
      Left            =   3240
      TabIndex        =   4
      Top             =   840
      Width           =   2820
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HanWorks"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   4320
      TabIndex        =   3
      Top             =   2520
      Width           =   1200
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Intializing..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Gathering Information..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3255
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
ProgressBar1.Value = 0
   With Timer1
   .Enabled = True
   .Interval = 30
   End With
Horizontal Me, &HFFFF00, &HFFC0C0
End Sub

Private Sub Form_Initialize()
' Show the ProgressBar and enable the timer.
   ProgressBar1.Visible = True
   Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Load Pic
Pic.Show
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1

If ProgressBar1.Value = 11 Then
  Unload Me
End If
End Sub
