VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmdraw 
   Caption         =   "Draw Fill Pad"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   3615
   Icon            =   "frmdraw.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   255
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   241
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdc 
      Left            =   2520
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2640
      ScaleHeight     =   465
      ScaleWidth      =   585
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   635
      ButtonWidth     =   1508
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Brush"
            Key             =   "brush"
            ImageIndex      =   1
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fill"
            Key             =   "fill"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Colour"
            Key             =   "colour"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Erase"
            Key             =   "erase"
            ImageIndex      =   2
            Style           =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   0
      ScaleHeight     =   215
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   151
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdraw.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdraw.frx":214E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuinsert 
         Caption         =   "Insert into Picture Viewer"
      End
      Begin VB.Menu sp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmdraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ExtFloodFill Lib "Gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
'variables used for drawing and filling
Dim X As Integer
Dim X1, Y1
Dim draw
Dim temp

Private Sub Form_Load()
X = 3
Picture2.BackColor = vbBlack
ImageList1.ListImages.Add X, , Picture2.Image
Toolbar1.Buttons("colour").Image = X
End Sub

Private Sub mnuclose_Click()
Unload Me
End Sub

Private Sub mnuinsert_Click()
With Pic
    .Image1.Stretch = False
    .Image1.Picture = Picture1.Image
    .Text1.Text = ""
    .Image2.Picture = .Image1.Picture
    .lblsize.Caption = "File Size:"
End With
Call Pic.actualsize
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'enable drawing
draw = 1
'if left mouse button, set up the line drawing
If Toolbar1.Buttons("brush").Value = tbrPressed Then
 X1 = X
 Y1 = Y
 Picture1.ForeColor = Picture2.BackColor
ElseIf Toolbar1.Buttons("erase").Value = tbrPressed Then
 X1 = X
 Y1 = Y
 Picture1.ForeColor = vbWhite
ElseIf Toolbar1.Buttons("fill").Value = tbrPressed Then
 Picture1.FillColor = Picture2.BackColor
 ExtFloodFill Picture1.hdc, X, Y, Picture1.Point(X, Y), 1
 DoEvents
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If draw = 1 Then
Picture1.Line (X1, Y1)-(X, Y)
X1 = X
Y1 = Y
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
draw = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Key

Case "brush"
 Toolbar1.Buttons("brush").Value = tbrPressed
 Toolbar1.Buttons("fill").Value = tbrUnpressed
 Toolbar1.Buttons("erase").Value = tbrUnpressed
Case "fill"
 Toolbar1.Buttons("brush").Value = tbrUnpressed
 Toolbar1.Buttons("fill").Value = tbrPressed
 Toolbar1.Buttons("erase").Value = tbrUnpressed
Case "colour"
 On Error GoTo Errhandler
 cdc.CancelError = True
 cdc.Flags = cdlCCRGBInit
 cdc.ShowColor
 Picture2.BackColor = cdc.Color
  X = X + 1
 ImageList1.ListImages.Add X, , Picture2.Image
 Toolbar1.Buttons("colour").Image = X
 Exit Sub

Errhandler:
   Exit Sub
Case "erase"
 Toolbar1.Buttons("brush").Value = tbrUnpressed
 Toolbar1.Buttons("fill").Value = tbrUnpressed
 Toolbar1.Buttons("erase").Value = tbrPressed
End Select

End Sub
