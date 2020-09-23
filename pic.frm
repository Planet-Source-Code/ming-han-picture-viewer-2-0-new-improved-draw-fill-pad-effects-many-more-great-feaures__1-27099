VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form pic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Picture Viewer"
   ClientHeight    =   7455
   ClientLeft      =   1380
   ClientTop       =   765
   ClientWidth     =   10920
   Icon            =   "pic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   728
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdeffect 
      BackColor       =   &H0000FF00&
      Caption         =   "Effects"
      Height          =   975
      Left            =   120
      MaskColor       =   &H00FF00FF&
      Picture         =   "pic.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.Frame border 
      Height          =   1215
      Index           =   0
      Left            =   960
      TabIndex        =   32
      Top             =   6240
      Width           =   135
   End
   Begin pictureviewer.FolderView fv1 
      Height          =   3375
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5953
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmdRotate 
      BackColor       =   &H0000FF00&
      Caption         =   "Rotate"
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdflip 
      BackColor       =   &H0000FF00&
      Caption         =   "Flip"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7080
      Width           =   1215
   End
   Begin MSComCtl2.UpDown cmdUpDown 
      Height          =   255
      Left            =   4920
      TabIndex        =   30
      Top             =   7080
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   450
      _Version        =   393216
      Value           =   90
      Max             =   359
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin VB.Frame border 
      Height          =   1215
      Index           =   2
      Left            =   6720
      TabIndex        =   23
      Top             =   6240
      Width           =   135
   End
   Begin VB.Frame RotateF 
      Height          =   495
      Left            =   3000
      TabIndex        =   24
      Top             =   6480
      Width           =   3735
      Begin VB.OptionButton ClockO 
         Caption         =   "Rotate Clockwise"
         Height          =   252
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton AntiClockO 
         Caption         =   "Rotate Anti-Clockwise"
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame border 
      Height          =   1215
      Index           =   1
      Left            =   2880
      TabIndex        =   22
      Top             =   6240
      Width           =   135
   End
   Begin VB.OptionButton HFlipO 
      Caption         =   "Flip Horizontally"
      Height          =   252
      Left            =   1320
      TabIndex        =   20
      Top             =   6600
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton VFlipO 
      Caption         =   "Flip Vertically"
      Height          =   252
      Left            =   1320
      TabIndex        =   19
      Top             =   6840
      Width           =   1215
   End
   Begin VB.PictureBox npb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3840
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   17
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox opb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   5880
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   16
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComCtl2.FlatScrollBar VScroll1 
      Height          =   5295
      Left            =   10560
      TabIndex        =   15
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   9340
      _Version        =   393216
      Appearance      =   2
      Max             =   100
      Orientation     =   8323072
   End
   Begin MSComCtl2.FlatScrollBar HScroll1 
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   5400
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   2
      Arrows          =   65536
      Max             =   100
      Orientation     =   8323073
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      TabIndex        =   4
      ToolTipText     =   "Shows Picture Path"
      Top             =   5880
      Width           =   6855
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pic.frx":30EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pic.frx":3ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pic.frx":4420
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pic.frx":5074
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pic.frx":5CC8
            Key             =   "out"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pic.frx":5E2C
            Key             =   "in"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   8280
      TabIndex        =   2
      Top             =   5760
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "go"
            Object.ToolTipText     =   "Go"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "actual"
            Object.ToolTipText     =   "Actual Size"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "in"
            Object.ToolTipText     =   "Zoom in"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "out"
            Object.ToolTipText     =   "Zoom out"
            ImageIndex      =   4
            Style           =   1
            Value           =   1
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "pic.frx":5F90
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   0
      OLEDropMode     =   1  'Manual
      System          =   -1  'True
      TabIndex        =   0
      Top             =   3360
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   3480
      ScaleHeight     =   351
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   463
      TabIndex        =   1
      Top             =   0
      Width           =   6975
      Begin MSComDlg.CommonDialog dlgopenfile 
         Left            =   0
         Top             =   3480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   135
         Left            =   0
         MousePointer    =   99  'Custom
         OLEDropMode     =   1  'Manual
         Stretch         =   -1  'True
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   975
      Left            =   6840
      ScaleHeight     =   915
      ScaleWidth      =   3915
      TabIndex        =   5
      Top             =   6480
      Width           =   3975
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   240
         Left            =   2400
         TabIndex        =   12
         Top             =   600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   423
         _Version        =   393216
         Value           =   5
         Max             =   15
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin VB.Frame Frame1 
         Caption         =   "C&apture"
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   2775
         Begin VB.OptionButton Option2 
            Caption         =   "active window"
            Height          =   255
            Left            =   1320
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "entire screen"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdok 
         BackColor       =   &H0000FF00&
         Caption         =   "&Capture Screen"
         Height          =   735
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   2040
         TabIndex        =   8
         Top             =   600
         Width           =   105
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Delay (in seconds):   "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1860
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   1080
      TabIndex        =   21
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label DegreeL 
      BackColor       =   &H00FFFFC0&
      Caption         =   "90 degrees"
      Height          =   255
      Left            =   3600
      TabIndex        =   28
      ToolTipText     =   "The Amount in Degrees that the Pictue will be Rotated"
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "by :"
      Height          =   255
      Left            =   3240
      TabIndex        =   29
      Top             =   7080
      Width           =   375
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   5160
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   552
      Y1              =   424
      Y2              =   424
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   2
      X1              =   0
      X2              =   552
      Y1              =   424
      Y2              =   424
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   10560
      Top             =   5400
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   0
      X1              =   224
      X2              =   224
      Y1              =   0
      Y2              =   376
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Picture Path:  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5880
      Width           =   1530
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   1
      X1              =   224
      X2              =   224
      Y1              =   0
      Y2              =   376
   End
   Begin VB.Label lblsize 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "File Size: "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   5640
      Width           =   3255
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnusave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuprint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnudraw 
         Caption         =   "&Draw Fill Pad"
         Shortcut        =   ^D
      End
      Begin VB.Menu sp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuwallpaper 
         Caption         =   "S&et as wallpaper"
      End
      Begin VB.Menu mnufreshdir 
         Caption         =   "&Refresh Directory"
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnupb 
      Caption         =   "Picture &Box"
      Begin VB.Menu mnucut 
         Caption         =   "Cu&t picture"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuco 
         Caption         =   "&Copy picture"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnupas 
         Caption         =   "&Paste picture"
         Shortcut        =   ^P
      End
      Begin VB.Menu sp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuapp 
         Caption         =   "&Appearance"
         Begin VB.Menu mnud 
            Caption         =   "3D"
         End
         Begin VB.Menu mnuflat 
            Caption         =   "flat"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuclear 
         Caption         =   "C&lear"
      End
   End
   Begin VB.Menu mnueffects 
      Caption         =   "&Effects"
      Begin VB.Menu mnubright 
         Caption         =   "Brighter"
      End
      Begin VB.Menu mnudark 
         Caption         =   "Darker"
      End
      Begin VB.Menu mnuneg 
         Caption         =   "Negative"
      End
      Begin VB.Menu mnblur 
         Caption         =   "Blur"
      End
   End
End
Attribute VB_Name = "pic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------
'/ Created by Teh Ming Han                          /
'/ E-mail: teh_minghan@hotmail.com                  /
'/ Singapore                                        /
'/ 7 September 2001                                 /
'/                                                  /
'/ REMEMBER TO RATE THIS AS-->EXCELLENT             /
'/                                                  /
'/ www.planet-source-code.com/vb/                   /
'/                                                  /
'/ Watch out for Picture Viewer 2.                  /
'----------------------------------------------------
Dim NewXpixel, NewYpixel, Xpixel, Ypixel, R1, G1, B1 As Integer
'--------------------------------------------------
Dim X As Long
Dim Y As Long
Const PI = 3.14159265358979
'------^-----^^-> flip & rotate--------------

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Const SPIF_UPDATEINIFILE = &H1
Const SPI_SETDESKWALLPAPER = 20
Const SPIF_SENDWININICHANGE = &H2
'------^-----^^-> wallpaper--------------

Dim TX As Long
Dim TY As Long
Dim ZoomDepth As Long
Dim Msg As Integer
Private Type PALETTEENTRY
   peRed As Byte
   peGreen As Byte
   peBlue As Byte
   peFlags As Byte
End Type

Option Base 0

Private Type LOGPALETTE
   palVersion As Integer
   palNumEntries As Integer
   palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
End Type

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function CreateCompatibleDC Lib "Gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "Gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "Gdi32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "Gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "Gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "Gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "Gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "Gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SelectPalette Lib "Gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "Gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Type PicBmp
   SIZE As Long
   Type As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Private Sub cmdeffect_Click()
PopupMenu mnueffects, , 8, 432
End Sub

Private Sub cmdflip_Click()
h_w_pic

MousePointer = vbHourglass
npb.Cls
If HFlipO Then
  For X = 0 To opb.Width
  For Y = 0 To opb.Height
    SetPixel npb.hdc, opb.Width - X, Y, GetPixel(opb.hdc, X, Y)
  Next
  npb.Picture = npb.Image
  Image2 = npb
  DoEvents
  Next
Else
  For X = 0 To opb.Width
  For Y = 0 To opb.Height
    SetPixel npb.hdc, X, opb.Height - Y, GetPixel(opb.hdc, X, Y)
  Next
  npb.Picture = npb.Image
  Image2 = npb
  DoEvents
  Next
End If

MousePointer = vbDefault
Image1 = Image2

End Sub

Private Sub cmdOK_Click()
Image1.Stretch = False
Dim EndTime As Date
Me.Visible = False
If Option2.Value = True Then
EndTime = DateAdd("s", Label3.Caption, Now)
Do Until Now > EndTime
DoEvents
Loop
Image1.Picture = CaptureActiveWindow()
Else
EndTime = DateAdd("s", Label3.Caption, Now)
Do Until Now > EndTime
DoEvents
Loop
Set Image1.Picture = CaptureScreen()
End If
Me.Visible = True
Me.SetFocus
click_refresh_paste
noscroll
Text1.Text = ""
Image2.Picture = Image1.Picture
File1.Refresh
lblsize.Caption = "File Size:"
End Sub

Private Sub size_ch(filesize)
mysize = filesize
If mysize >= 1000000 Then
  mysize = Fix(mysize / 1000)
  mysize = mysize / 1000
  lblsize.Caption = "File Size: " & mysize & " mb"
ElseIf mysize >= 1000 Then
  mysize = mysize / 1000
  lblsize.Caption = "File Size: " & mysize & " kb"
Else
  lblsize.Caption = "File Size: " & mysize & " bytes"
End If
End Sub

Private Sub cmdRotate_Click()
h_w_pic

Dim a As Double
Dim SinA As Double
Dim CosA As Double
Dim cx As Long
Dim cy As Long

MousePointer = vbHourglass

npb.Picture = Nothing
Image2 = Nothing

cx = opb.Width \ 2
cy = opb.Height \ 2

If ClockO Then
  a = cmdUpDown.Value
Else
  a = 360 - cmdUpDown.Value
End If

a = a / 180 * PI

SinA = Sin(a)
CosA = Cos(a)

For X = -cx To cx
For Y = -cy To cy
  SetPixel npb.hdc, (X * CosA) - (Y * SinA) + cx, (X * SinA) + (Y * CosA) + cy, GetPixel(opb.hdc, X + cx, Y + cy)
Next
npb.Picture = npb.Image
Image2 = npb
DoEvents
Next

MousePointer = vbDefault
Image1 = Image2

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim GotoVal
Dim Gointo

    GotoVal = Me.Height / 2
    For Gointo = 1 To GotoVal
        DoEvents
        Me.Height = Me.Height - 100
        Me.Top = (Screen.Height - Me.Height) \ 2
        If Me.Height <= 500 Then Exit For
    Next Gointo

    Me.Height = 30
    GotoVal = Me.Width / 2
    For Gointo = 1 To GotoVal
        DoEvents
        Me.Width = Me.Width - 100
        Me.Left = (Screen.Width - Me.Width) \ 2
        If Me.Width <= 2000 Then Exit For
    Next Gointo
End Sub

Private Sub fv1_FolderClick(FolderName As String, FullPath As String)
On Error Resume Next
File1.Path = FullPath
Text1.Text = FullPath
End Sub

Private Sub effect_sub(effect As String)

MousePointer = vbHourglass
h_w_pic

Dim dloop As Integer
 If opb.Height >= opb.Width Then
 dloop = opb.Height
Else
 dloop = opb.Width
End If

npb.Cls
Ypixel = 0
Do
If Xpixel <= dloop Then
R1 = (opb.Point(Xpixel, Ypixel) And &HFF&) 'Get Red Value
G1 = ((opb.Point(Xpixel, Ypixel) And &HFF00&) / &H100&) 'Get Green Value
B1 = ((opb.Point(Xpixel, Ypixel) And &HFF0000) / &H10000) 'Get Blue Value

If effect = "negative" Then
 R1 = (255 - R1)
 G1 = (255 - G1)
 B1 = (255 - B1)
ElseIf effect = "bright" Then
 R1 = (R1 + 50)
 G1 = (G1 + 50)
 B1 = (B1 + 50)

ElseIf effect = "dark" Then
 R1 = (R1 - 50)
 G1 = (G1 - 50)
 B1 = (B1 - 50)

ElseIf effect = "blur" Then
 Randomize
 If Xpixel Mod 2 = 1 Then
 R1 = R1 + Int(Rnd * 255)
 G1 = G1 + Int(Rnd * 255)
 B1 = B1 + Int(Rnd * 255)
 Else
 R1 = R1 - Int(Rnd * 255)
 G1 = G1 - Int(Rnd * 255)
 B1 = B1 - Int(Rnd * 255)
 End If

End If
  If R1 < 0 Then R1 = 0
  If B1 < 0 Then B1 = 0
  If G1 < 0 Then G1 = 0
  If R1 > 255 Then R1 = 255
  If B1 > 255 Then B1 = 255
  If G1 > 255 Then G1 = 255
 npb.PSet (Xpixel, Ypixel), RGB(R1, G1, B1)
Xpixel = (Xpixel + 1)
Else
Ypixel = (Ypixel + 1)
Xpixel = 0
End If
Loop Until Ypixel = dloop
Image1.Picture = Nothing
Image2.Picture = npb.Image
Image1.Picture = npb.Image
MousePointer = vbDefault

End Sub

Private Sub Image1_Click()
 On Error GoTo BadZoom

If Toolbar1.Buttons("in").Value = tbrPressed Then
 
           If ZoomDepth >= 10 Then Beep: Exit Sub
           'Notice the "Image1.Width / 4" that is used here. This merely
           'increases the image by 25%. You may use a different number
           'than "4" to change your zoom ratio, but make sure you use
           'the same number through your code.
           Image1.Width = Image1.Width + (Image1.Width / 4)
           Image1.Height = Image1.Height + (Image1.Height / 4)

           If Image1.Width < Picture1.Width Then
               Image1.Left = 0
           Else
               'Else, everything seems to be good
               'so we will zoom in as calculated below.
               'NOTICE that this is where we maintain
               'our "point of view". What I mean is,
               'our mouse cursor is pointed at a specific
               'area of the image, so when we zoom in, we
               'want to see that same area at a closer view.
               'The "X" in the code, directly below, is part
               'of the calculation of the horizontal mouse
               'positio, which in turn sets the scroll bar
               'properly. Thus the image has shifted the
               'correct amount.
               Set_Scrolls
               
               If HScroll1.Value + ((X / TX) / 4) > HScroll1.Max Then
                   'This "IF" statement makes sure that our scroll value
                   'does not exceed our Scroll MAX when zooming
                   'in near the far right of the image. If it does
                   'exceed, we will use the maximum scroll value
                   HScroll1.Value = HScroll1.Max
               Else
                   HScroll1.Value = HScroll1.Value + ((X / TX) / 4)
               End If

           End If

           'The "IF" statement below is the same
           'as the one above, but it will now refer to the
           'image height instead of the width

           If Image1.Height < Picture1.Height Then
           Else
               Set_Scrolls

               If VScroll1.Value + ((Y / TY) / 4) > VScroll1.Max Then
                   VScroll1.Value = VScroll1.Max
               Else
                   VScroll1.Value = VScroll1.Value + ((Y / TY) / 4)
               End If

           End If

           ZoomDepth = ZoomDepth + 1 'To keep track of how many times we soomed in

ElseIf Toolbar1.Buttons("out").Value = tbrPressed Then
If Image1.Width <= 10 Then Beep: Exit Sub
           If Image1.Height <= 10 Then Beep: Exit Sub
           Image1.Width = Image1.Width - (Image1.Width / 4)
           Image1.Height = Image1.Height - (Image1.Height / 4)

           If Image1.Width < Picture1.Width Then
               'Do nothing
           Else

               If HScroll1.Value - ((X / TX) / 4) > HScroll1.Max Then
                   HScroll1.Value = HScroll1.Max
               ElseIf HScroll1.Value - ((X / TX) / 4) < 1 Then
                   HScroll1.Value = 1
               Else
                   HScroll1.Value = HScroll1.Value - ((X / TX) / 4)
               End If

           End If

           If Image1.Height < Picture1.Height Then
               Image1.Top = 0
           Else

               If VScroll1.Value - ((Y / TY) / 4) > VScroll1.Max Then
                   VScroll1.Value = VScroll1.Max
               ElseIf VScroll1.Value - ((Y / TY) / 4) < 1 Then
                   VScroll1.Value = 1
               Else
                   VScroll1.Value = VScroll1.Value - ((Y / TY) / 4)
               End If

           End If

           ZoomDepth = ZoomDepth - 1 'Deduct each time we zoom out
       End If

       Set_Scrolls 'Jump to the "Set_Scrolls Sub" here
       'which will determine when to enable
       'or disable a scroll bar.
       Exit Sub
BadZoom:
       Resume Next
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Toolbar1.Buttons("in").Value = tbrPressed Then
Image1.MouseIcon = ImageList1.ListImages("in").Picture
ElseIf Toolbar1.Buttons("out").Value = tbrPressed Then
Image1.MouseIcon = ImageList1.ListImages("out").Picture
End If

End Sub

Private Sub mnblur_Click()
effect_sub "blur"
End Sub

Private Sub mnubright_Click()
effect_sub "bright"
End Sub

Private Sub mnuclear_Click()
Image1.Picture = Nothing
End Sub

Private Sub mnucut_Click()
mnuco_Click
Image1 = Nothing
End Sub

Private Sub mnudark_Click()
effect_sub "dark"
End Sub

Private Sub mnudraw_Click()
With frmdraw
    .Show
    .Picture1.Cls
    .Picture1.Picture = Image1.Picture
    .Width = (.Picture1.ScaleWidth + 8) * Screen.TwipsPerPixelX
    .Height = (.Picture1.ScaleHeight + 70) * Screen.TwipsPerPixelY
    .Refresh
End With
End Sub

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub File1_Click()
On Error Resume Next
    Text1.Text = File1.Path + "\" + File1.FileName
    Image1.Stretch = False
    Image1.Picture = LoadPicture(File1.Path + "\" + File1.FileName)
    click_refresh_paste
    noscroll
    
mysize = FileLen(File1.Path + "\" + File1.FileName)
size_ch (mysize)
End Sub

Private Sub Form_Load()
HScroll1.Enabled = False
VScroll1.Enabled = False
    TX = Screen.TwipsPerPixelX
    TY = Screen.TwipsPerPixelY
   File1.Pattern = "*.BMP;*.DIB;*.ICO;*.WMF;*.EMF;*.GIF;*.JPG;*.JPEG;*.CUR"
End Sub

Private Sub HScroll1_Change()
     HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
Image1.Left = -HScroll1.Value

If HScroll1.Value = 1 Then
HScroll1.Arrows = cc2RightDown
ElseIf HScroll1.Value = HScroll1.Max Then
HScroll1.Arrows = cc2LeftUp
Else
HScroll1.Arrows = cc2Both
End If
End Sub

Private Sub mnuco_Click()
On Error Resume Next
   Clipboard.Clear   ' Clear Clipboard.
   Clipboard.SetData Image1.Picture
End Sub

Private Sub mnud_Click()
Picture1.Appearance = 1
mnud.Checked = True
mnuflat.Checked = False
End Sub


Private Sub mnuflat_Click()
Picture1.Appearance = 0
mnud.Checked = False
mnuflat.Checked = True
End Sub

Private Sub mnufreshdir_Click()
fv1.Refresh
File1.Refresh
End Sub

Private Sub mnublur_Click()
effect_sub "blur"
End Sub

Private Sub mnuneg_Click()
effect_sub "negative"
End Sub

Private Sub mnupas_Click()
       Image1.Stretch = False
       Image1.Picture = Clipboard.GetData()
       Text1.Text = ""
       click_refresh_paste
       noscroll
       Image2.Picture = Clipboard.GetData()
End Sub

Private Sub mnuprint_Click()
If Error Then Exit Sub
Printer.PaintPicture Image1.Picture, 0, 0, Image1.Width, Image1.Height
    Printer.EndDoc
End Sub

Private Sub mnusave_Click()
With dlgopenfile
      .DefaultExt = ".bmp"
      .Filter = "Bitmap Image (*.bmp)|*.bmp"
End With

dlgopenfile.CancelError = True
On Error GoTo Errhandler
dlgopenfile.Flags = cdlOFNOverwritePrompt
dlgopenfile.ShowSave
SavePicture Image1.Picture, dlgopenfile.FileName
Msg = MsgBox("Picture saved!", vbInformation)
Errhandler:
   ' User pressed Cancel button.
   Exit Sub
End Sub

Private Sub mnuwallpaper_Click()
SavePicture Image1.Picture, "c:\windows\pvwallpaper"
X = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, "c:\windows\pvwallpaper", SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Key

Case "go"
 Image1.Stretch = False
 Image1.Picture = LoadPicture(Text1.Text)
 click_refresh_paste
 noscroll
Case "actual"
 actualsize
Case "in"
 If Toolbar1.Buttons("in").Value = tbrUnpressed Then
  Toolbar1.Buttons("out").Value = tbrPressed
 Else
  Toolbar1.Buttons("out").Value = tbrUnpressed
 End If
Case "out"
 If Toolbar1.Buttons("out").Value = tbrUnpressed Then
  Toolbar1.Buttons("in").Value = tbrPressed
 Else
  Toolbar1.Buttons("in").Value = tbrUnpressed
 End If
End Select

End Sub

Private Sub UpDown1_Change()
Label3.Caption = UpDown1.Value
End Sub

Private Sub VScroll1_Change()
VScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()
Image1.Top = -VScroll1.Value

If VScroll1.Value = 1 Then
VScroll1.Arrows = cc2RightDown
ElseIf VScroll1.Value = VScroll1.Max Then
VScroll1.Arrows = cc2LeftUp
Else
VScroll1.Arrows = cc2Both
End If
End Sub
 Public Sub Set_Scrolls()
       If Image1.Width > Picture1.Width Then
           'If the Image is wider than the picture box
           'then we want to enable the horizontal
           'scroll bar
           HScroll1.Enabled = True
           HScroll1.Value = 1
           'We can set the maximum vale for this scroll bar
           'based on the difference in the width of
           'the picture box and the image.
           'Note: since our scale mode is PIXEL and not TWIP
           'we have a scroll bar that is efficient. TWIPs could
           'very easily make the scroll bar MAX value
           'into the thousands.
           HScroll1.Max = Image1.Width - Picture1.Width
           HScroll1.Min = 1
           HScroll1.SmallChange = 1
           HScroll1.LargeChange = 8
       Else
           'Else, our image is not wider than the picture box and
           'we just dissable the scroll bar.
           HScroll1.Enabled = False
       End If
       'The code below is the same as the above code, but
       'deals with the height.
       If Image1.Height > Picture1.Height Then
           VScroll1.Enabled = True
           VScroll1.Value = 1
           VScroll1.Max = Image1.Height - Picture1.Height
           VScroll1.Min = 1
           VScroll1.SmallChange = 1
           VScroll1.LargeChange = 8
       Else
           VScroll1.Enabled = False
       End If

   End Sub

Private Sub click_refresh_paste()
Image1.Stretch = True
      
Image1.Enabled = True
Set_Scrolls

HScroll1.Value = 1
VScroll1.Value = 1
ZoomDepth = 0
lblsize.Caption = "File Size:"
End Sub

Private Sub noscroll()
Toolbar1.Buttons("out").Value = tbrPressed
Toolbar1.Buttons("in").Value = tbrUnpressed
If VScroll1.Enabled = True Then
Do Until VScroll1.Enabled = False
Image1_Click
Loop
End If
       
If HScroll1.Enabled = True Then
Do Until HScroll1.Enabled = False
Image1_Click
Loop
End If
End Sub

Public Function CaptureScreen() As Picture
 Dim hWndScreen As Long
   hWndScreen = GetDesktopWindow()
   Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
End Function
Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture

  Dim hDCMemory As Long
  Dim hBmp As Long
  Dim hBmpPrev As Long
  Dim R As Long
  Dim hDCSrc As Long
  Dim hPal As Long
  Dim hPalPrev As Long
  Dim RasterCapsScrn As Long
  Dim HasPaletteScrn As Long
  Dim PaletteSizeScrn As Long
  Dim LogPal As LOGPALETTE

   If Client Then
      hDCSrc = GetDC(hWndSrc) ' Get device context for client area.
   Else
      hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire
                                    ' window.
   End If

   hDCMemory = CreateCompatibleDC(hDCSrc)
   ' Create a bitmap and place it in the memory DC.
   hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
   hBmpPrev = SelectObject(hDCMemory, hBmp)

   ' Get screen properties.
   RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
                                                      ' capabilities.
   HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
                                                        ' support.
   PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
                                                        ' palette.

   ' If the screen has a palette make a copy and realize it.
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      ' Create a copy of the system palette.
      LogPal.palVersion = &H300
      LogPal.palNumEntries = 256
      R = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
      hPal = CreatePalette(LogPal)
      ' Select the new palette into the memory DC and realize it.
      hPalPrev = SelectPalette(hDCMemory, hPal, 0)
      R = RealizePalette(hDCMemory)
   End If

   ' Copy the on-screen image into the memory DC.
   R = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)

' Remove the new copy of the  on-screen image.
   hBmp = SelectObject(hDCMemory, hBmpPrev)

   ' If the screen has a palette get back the palette that was
   ' selected in previously.
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      hPal = SelectPalette(hDCMemory, hPalPrev, 0)
   End If

   ' Release the device context resources back to the system.
   R = DeleteDC(hDCMemory)
   R = ReleaseDC(hWndSrc, hDCSrc)

   ' Call CreateBitmapPicture to create a picture object from the
   ' bitmap and palette handles. Then return the resulting picture
   ' object.
   Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
  Dim R As Long

   Dim Pic As PicBmp
   ' IPicture requires a reference to "Standard OLE Types."
   Dim IPic As IPicture
   Dim IID_IDispatch As GUID

   With IID_IDispatch
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With

   With Pic
      .SIZE = Len(Pic)          ' Length of structure.
      .Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
      .hBmp = hBmp              ' Handle to bitmap.
      .hPal = hPal              ' Handle to palette (may be null).
   End With

   ' Create Picture object.
   R = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

   ' Return the new Picture object.
   Set CreateBitmapPicture = IPic
End Function
Public Function CaptureActiveWindow() As Picture

    Dim hWndActive As Long
    Dim R As Long
    Dim RectActive As RECT
    
    hWndActive = GetForegroundWindow()
    
    R = GetWindowRect(hWndActive, RectActive)
    
    ' Call CaptureWindow to capture the active window given its
    ' handle and return the Resulting Picture object.
    Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, RectActive.Right - RectActive.Left, RectActive.Bottom - RectActive.Top)
End Function
Public Sub actualsize()
Image1.Stretch = False
 If Text1.Text = "" Then
 Image1.Picture = Image2.Picture
 click_refresh_paste
 Else
 Image1.Picture = LoadPicture(File1.Path + "\" + File1.FileName)
 click_refresh_paste
 End If
End Sub

Private Sub cmdUpDown_Change()
DegreeL = cmdUpDown.Value & " Degrees"
End Sub

Private Function h_w_pic()
lblsize.Caption = "File Size:"
Text1.Text = ""
opb.Height = Image1.Height
opb.Width = Image1.Width
Image2.Height = Image1.Height
Image2.Width = Image1.Width
opb = Image1
Image2.Picture = LoadPicture()
npb = LoadPicture()
npb.Width = opb.Width
npb.Height = opb.Height

End Function
