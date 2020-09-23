VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl FolderView 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2760
   ScaleHeight     =   3690
   ScaleWidth      =   2760
   ToolboxBitmap   =   "FolderTreeView.ctx":0000
   Begin MSComctlLib.TreeView tvDirs 
      Height          =   2745
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   4842
      _Version        =   393217
      Indentation     =   88
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imlDir"
      Appearance      =   0
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList imlDir 
      Left            =   1680
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderTreeView.ctx":0312
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderTreeView.ctx":0BEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderTreeView.ctx":14CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderTreeView.ctx":1DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderTreeView.ctx":2682
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderTreeView.ctx":29D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderTreeView.ctx":2ACA
            Key             =   "DragCopy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderTreeView.ctx":2C2E
            Key             =   "DragMove"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FolderView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************
' Abstract FolderView Control
' http://abstractvb.com/
' Copyright abstractvb.com, 2000. All Rights Reserved.
'*********************************************************

Option Explicit

Private nodCurrent As Node
Private fso As New FileSystemObject

Event FolderDoubleClick(FolderName As String, FullPath As String)
Event FolderClick(FolderName As String, FullPath As String)
Attribute FolderClick.VB_MemberFlags = "200"

Event Collapse(ByVal Node As Node)
Event Expand(ByVal Node As Node)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode Then
        GetDrives
    End If
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next

    tvDirs.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Public Sub GetDrives()
    Dim fldr As Scripting.folder
    Dim dr As Scripting.Drive
    Dim RootNode As Node
    Dim drNode As Node
    Dim foNode As Node
        
    On Error Resume Next
    
    tvDirs.Nodes.Clear
        
    For Each dr In fso.Drives
        Set drNode = tvDirs.Nodes.Add(, , , dr.Path, dr.DriveType)
        drNode.Tag = "DRIVE"
        drNode.Sorted = True
        
        '*****************************************************
        ' Create a Dummy Node so we get a + beside each drive
        '*****************************************************
        Set foNode = tvDirs.Nodes.Add(drNode, tvwChild)
    Next
End Sub

Private Sub tvDirs_DblClick()
    On Error Resume Next
    
    Dim lsPath As String
        
    If Not nodCurrent Is Nothing Then
        lsPath = nodCurrent.FullPath
        If Right$(lsPath, 1) <> "\" Then lsPath = lsPath & "\"
        RaiseEvent FolderDoubleClick(nodCurrent.Text, lsPath)
    End If
End Sub

Private Sub tvDirs_Expand(ByVal Node As MSComctlLib.Node)
    RaiseEvent Expand(Node)
    Dim fldr As Scripting.folder
    Dim rfldr As Scripting.folder
    Dim lsPath As String
    
    lsPath = Node.FullPath
    If Right(lsPath, 1) <> "\" Then lsPath = lsPath & "\"
    
    If Not fso.FolderExists(lsPath) = False Then
        Set rfldr = fso.GetFolder(lsPath)
        UserControl.MousePointer = vbHourglass
    
        RemoveChildren Node
        
        For Each fldr In rfldr.SubFolders
            AddSubFolder Node, fldr
        Next
        UserControl.MousePointer = vbDefault
    Else
        MsgBox "Disk not ready!", vbCritical, App.Title
    End If
End Sub

'*****************************************************
'Recursively adds subfolders
'*****************************************************
Public Sub AddSubFolder(ByVal parent As Node, ByVal folder As Scripting.folder)
    Dim foNode As Node
    Dim tmpNode As Node
    Dim fo As Scripting.folder
    
    On Error Resume Next
         
    Set foNode = tvDirs.Nodes.Add(parent, tvwChild, , folder.Name, 6)
    foNode.ExpandedImage = 5
    foNode.Sorted = True
    
    'Create Dummy Node
    If folder.SubFolders.Count > 0 Then
        Set tmpNode = tvDirs.Nodes.Add(foNode, tvwChild)
    End If
End Sub

'*****************************************************
'Removes all the child nodes from the node passed
'*****************************************************
Private Sub RemoveChildren(ByVal nodx As Node)
    Dim n As Node
        
    Do While nodx.Children > 0
        Set n = nodx.Child
        tvDirs.Nodes.Remove n.Index
    Loop
End Sub

Private Sub tvDirs_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    Dim lsPath As String
    
    Set nodCurrent = Node
        
    If Not nodCurrent Is Nothing Then
        lsPath = nodCurrent.FullPath
        If Right$(lsPath, 1) <> "\" Then lsPath = lsPath & "\"
        RaiseEvent FolderClick(Node.Text, lsPath)
    End If
End Sub

Private Sub tvDirs_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    Dim fso As New Scripting.FileSystemObject
    Dim f As Scripting.File
    Dim fldr As Scripting.folder
    Dim i As Long
    Dim lsFilestring As String
    Dim lsPath As String
    
    If Not tvDirs.SelectedItem Is Nothing Then
        AllowedEffects = vbDropEffectCopy
        Data.Clear
                
        lsFilestring = ""
        
        lsPath = tvDirs.SelectedItem.FullPath
        
        If Right$(lsPath, 1) <> "\" Then lsPath = lsPath & "\"
        
        Set fldr = fso.GetFolder(lsPath)
                
        For Each f In fldr.Files
            If UCase$(f.Name) Like "*.GIF" Or UCase$(f.Name) Like "*.JPG" Or UCase$(f.Name) Like "*.BMP" Then
                If lsFilestring = "" Then
                    lsFilestring = f.Path
                Else
                    lsFilestring = lsFilestring & "," & f.Path
                End If
            End If
        Next
              
        Data.SetData lsFilestring, vbCFText
    End If
End Sub

Private Sub UserControl_Terminate()
    Set fso = Nothing
End Sub

Public Property Get BorderStyle() As MSComctlLib.BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As MSComctlLib.BorderStyleConstants)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub tvDirs_Collapse(ByVal Node As Node)
    RaiseEvent Collapse(Node)
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub tvDirs_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set tvDirs.SelectedItem = tvDirs.HitTest(x, y)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub tvDirs_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    GetDrives
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

