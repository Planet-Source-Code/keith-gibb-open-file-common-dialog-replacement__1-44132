VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Comm1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Common Dialog Replacement"
   ClientHeight    =   7335
   ClientLeft      =   1950
   ClientTop       =   1590
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   10590
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   7080
      TabIndex        =   24
      Text            =   "Text2"
      Top             =   420
      Width           =   2115
   End
   Begin VB.PictureBox pbxPrev 
      Height          =   2895
      Left            =   240
      ScaleHeight     =   2835
      ScaleWidth      =   5655
      TabIndex        =   20
      Top             =   660
      Visible         =   0   'False
      Width           =   5715
      Begin VB.CommandButton cmdPrev 
         Height          =   255
         Left            =   4785
         Picture         =   "Form.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Previous Preview"
         Top             =   2565
         Width           =   255
      End
      Begin VB.CommandButton cmdNext 
         Height          =   255
         Left            =   5085
         Picture         =   "Form.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Next Preview"
         Top             =   2565
         Width           =   255
      End
      Begin VB.CommandButton cmdPrevCancel 
         Height          =   255
         Left            =   5385
         Picture         =   "Form.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Close Preview Window"
         Top             =   2565
         Width           =   255
      End
      Begin VB.Label Label13 
         Caption         =   "File size:"
         ForeColor       =   &H80000015&
         Height          =   195
         Left            =   3360
         TabIndex        =   32
         Top             =   1620
         Width           =   1635
      End
      Begin VB.Label Label12 
         Caption         =   "Date/Time:"
         ForeColor       =   &H80000015&
         Height          =   195
         Left            =   3360
         TabIndex        =   31
         Top             =   840
         Width           =   1635
      End
      Begin VB.Label Label11 
         Caption         =   "File name:"
         ForeColor       =   &H80000015&
         Height          =   195
         Left            =   3360
         TabIndex        =   30
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   255
         Left            =   3360
         TabIndex        =   29
         Top             =   2070
         Width           =   1635
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   435
         Left            =   3360
         TabIndex        =   28
         Top             =   1860
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   435
         Left            =   3360
         TabIndex        =   27
         Top             =   465
         Width           =   2235
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   435
         Left            =   3360
         TabIndex        =   26
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         Caption         =   "Preview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   0
         TabIndex        =   25
         Top             =   2535
         Width           =   5655
      End
      Begin VB.Image imgPrev 
         Height          =   915
         Left            =   480
         Stretch         =   -1  'True
         Top             =   660
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4860
      TabIndex        =   19
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtFileName 
      Height          =   315
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3720
      Width           =   3495
   End
   Begin VB.ComboBox cmbFileTypes 
      Height          =   315
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   4140
      Width           =   3495
   End
   Begin MSComctlLib.ImageList imlTools 
      Left            =   6360
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form.frx":03DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form.frx":0978
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form.frx":0F12
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form.frx":106C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4725
      ScaleHeight     =   330
      ScaleWidth      =   1260
      TabIndex        =   13
      Top             =   180
      Width           =   1260
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlTools"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Previous"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Create New Folder"
               ImageIndex      =   2
               Object.Width           =   500
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Large Icons"
               ImageIndex      =   4
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Text            =   "Small Icons"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Large Icons"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlDrives 
      Left            =   6360
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form.frx":11C6
            Key             =   "open folder"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo icmDrives 
      Height          =   330
      Left            =   240
      TabIndex        =   12
      Top             =   180
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      Text            =   "ImageCombo1"
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   8520
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
   End
   Begin VB.FileListBox File2 
      Height          =   1065
      Left            =   7080
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   7080
      TabIndex        =   6
      Top             =   1020
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   7080
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.PictureBox pic16 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1260
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   4860
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic32 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2220
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton cmdOKay 
      Caption         =   "Open"
      Height          =   315
      Left            =   4860
      TabIndex        =   1
      Top             =   4140
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7080
      TabIndex        =   0
      Text            =   "c:\"
      Top             =   60
      Width           =   2100
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   6360
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   6360
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   2895
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   660
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   5106
      View            =   2
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "File"
         Text            =   "File"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Path"
         Text            =   "Path"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Image imgHolder 
      Height          =   615
      Left            =   7740
      Top             =   4860
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "File name:"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   3780
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Files of type:"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   7080
      TabIndex        =   11
      Top             =   5460
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   7080
      TabIndex        =   9
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   7080
      TabIndex        =   8
      Top             =   4860
      Width           =   1155
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu mnuPreview 
         Caption         =   "Preview"
      End
   End
End
Attribute VB_Name = "Comm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------
'Replacement common dialogue
'© Keith Gibb (On Screen Media Ltd.), 2003
'gibbo912@aol.com
'If you use this, please credit me and the guys below :)
'--------------------------------------------------------------------
'Thanks to T De Lange (tomdl@attglobal.net) and Peter Meier
'for their icon extraction code
'Thanks to planet-source-code.com for making the impossible, possible
'--------------------------------------------------------------------


'Icon Sizes in pixels
Private Const LARGE_ICON As Integer = 32
Private Const SMALL_ICON As Integer = 16
Private Const MAX_PATH = 260

Private Const ILD_TRANSPARENT = &H1       'Display transparent

'ShellInfo Flags
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Private Const SHGFI_LARGEICON = &H0       'Large icon
Private Const SHGFI_SMALLICON = &H1       'Small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400

Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO                   'As required by ShInfo
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type


'----------------------------------------------------------
'Functions & Procedures
'----------------------------------------------------------
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" _
    (ByVal himl&, ByVal i&, ByVal hDCDest&, _
    ByVal X&, ByVal Y&, ByVal flags&) As Long


'----------------------------------------------------------
'Private variables
'----------------------------------------------------------
Private ShInfo As SHFILEINFO


Private Declare Function GetDriveType Lib "Kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetVolumeInformation Lib "Kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Dim FileExt As String
Dim CheckType As Integer
Dim FullFoldName As String
Dim FoldName As String

Private Const MAX_FILENAME_LEN = 256


Private Sub Display()
Me.MousePointer = 11
txtFileName = ""
Dir1.Path = Text1.Text
File2.Path = Text1.Text
lvFiles.ListItems.Clear
For seq = 0 To Dir1.ListCount - 1
  FullFoldName = Dir1.List(seq)
  For seq2 = Len(FullFoldName) To 1 Step -1
   If Mid(FullFoldName, seq2, 1) = "\" Then
    FoldName = Right(FullFoldName, Len(FullFoldName) - seq2)
    Exit For
   End If
  Next seq2
  lvFiles.ListItems.Add , FullFoldName, FoldName
Next seq
For seq = 0 To File2.ListCount - 1
 If Mid(File2.List(seq), Len(File2.List(seq)) - 3, 1) = "." Then
  lvFiles.ListItems.Add , Text1.Text & File2.List(seq), File2.List(seq)
 End If
Next seq
Call Initialise(lvFiles)
Call GetAllIcons(lvFiles)
Call ShowIcons(lvFiles)
Me.MousePointer = vbDefault
End Sub

Private Sub DisplayDrives()
Dim DriveName As String
Dim DriveVolName As String
Dim IconType As Integer

For seq = 0 To Drive1.ListCount - 1
 GetDriveIcons Left(Drive1.List(seq), 2) & "\", seq + 1, Left(Drive1.List(seq), 2) & "\"
Next seq
For seq = 0 To Drive1.ListCount - 1
    Select Case GetDriveType(Left(Drive1.List(seq), 2) & "\")
        Case 2
            IconType = 1
            If UCase(Left(Drive1.List(seq), 1)) = "A" Or UCase(Left(Drive1.List(seq), 1)) = "B" Then
             DriveName = "Floppy Disk"
            Else
             DriveName = "Removeable Disk"
            End If
        Case 3
            IconType = 2
            DriveName = "Local Disk"
        Case Is = 4
            IconType = 3
            DriveName = "Network Drive"
        Case Is = 5
            IconType = 4
            DriveName = "CD-Rom"
        Case Else
            IconType = 2
            DriveName = "Unrecognized"
    End Select
  Set icmDrives.ImageList = imlDrives
  icmDrives.ComboItems.Add , Left(Drive1.List(seq), 2) & "\", DriveName & " (" & UCase(Left(Drive1.List(seq), 2)) & ")  " & Right(Drive1.List(seq), Len(Drive1.List(seq)) - 2), Left(Drive1.List(seq), 2) & "\"
  
  If Left(Text1.Text, 2) = Left(Drive1.List(seq), 2) And Len(Text1.Text) > 3 Then
   Dim LastSlash As Integer
   Dim FoldIndent As Integer
   FoldIndent = 1
   LastSlash = 4
   For seq2 = 4 To Len(Text1.Text)
    If Mid(Text1.Text, seq2, 1) = "\" Then
     icmDrives.ComboItems.Add , Left(Text1.Text, Len(Text1.Text) - (Len(Text1.Text) - seq2)), Mid(Text1.Text, LastSlash, (seq2 - LastSlash)), "open folder", , FoldIndent
     LastSlash = seq2 + 1
     FoldIndent = FoldIndent + 1
    End If
   Next seq2
  End If

Next seq
 icmDrives.ComboItems(Text1.Text).Selected = True
End Sub




Private Sub GetDriveIcons(FileName As String, Index As Long, ImageName As String)
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
'On Error Resume Next
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
Dim r As Long


Label1.Caption = Index

'Get a handle to the small icon
hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'Get a handle to the large icon
hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

'If the handle(s) exists, load it into the picture box(es)
If hLIcon <> 0 Then
  'Large Icon
  With pic32
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hLIcon, ShInfo.iIcon, pic32.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  'Small Icon
  With pic16
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hSIcon, ShInfo.iIcon, pic16.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  Set imgObj = imlDrives.ListImages.Add(Index, ImageName, pic16.Image)
End If
End Sub

Private Sub cmbFileTypes_Click()
 Select Case cmbFileTypes.ListIndex
  Case 0
   File2.FileName = "*.*"
  Case 1
   File2.FileName = "*.swf"
  Case 2
   File2.FileName = "*.fla"
  Case 3
   File2.FileName = "*.swf;*.fla"
  Case 4
   File2.FileName = "*.bmp;*.jpg;*.wmf;*.gif;*.ico;*.cur"
  Case 5
   File2.FileName = "*.msk"
 End Select
 Call Display
End Sub


Private Sub Command2_Click()
End Sub







Private Sub Initialise(LV As Object)
'-----------------------------------------------
'Initialise the controls
'-----------------------------------------------
On Local Error Resume Next

'Break the link to iml lists
'lvFiles.ListItems.Clear
LV.Icons = Nothing
LV.SmallIcons = Nothing

'Clear the image lists
iml32.ListImages.Clear
iml16.ListImages.Clear

End Sub


Private Sub cmdNext_Click()
  lvFiles.ListItems(lvFiles.SelectedItem.Index).Selected = False
  For seq = lvFiles.SelectedItem.Index + 1 To lvFiles.ListItems.Count
   Select Case UCase(Right(lvFiles.ListItems(seq).Text, 4))
    Case ".BMP", ".JPG", ".ICO", ".GIF", ".CUR", ".WMF"
     lvFiles.ListItems(seq).Selected = True
     txtFileName = lvFiles.ListItems(seq).Text
     Call mnuPreview_Click
     Exit Sub
   End Select
  Next seq
End Sub

Private Sub cmdOKay_Click()
Dim i As Integer
Dim FileStr As String
 For i = 1 To lvFiles.ListItems.Count
   If lvFiles.ListItems(i).Selected = True Then
      If Mid(lvFiles.SelectedItem.Key, Len(lvFiles.SelectedItem.Key) - 3, 1) = "." Then
       FileStr = FileStr & lvFiles.ListItems(i).Key
      End If
   End If
 Next i
MsgBox (FileStr)
End Sub

Private Sub cmdPrev_Click()
  lvFiles.ListItems(lvFiles.SelectedItem.Index).Selected = False
  For seq = lvFiles.SelectedItem.Index - 1 To 1 Step -1
   Select Case UCase(Right(lvFiles.ListItems(seq).Text, 4))
    Case ".BMP", ".JPG", ".ICO", ".GIF", ".CUR", ".WMF"
     lvFiles.ListItems(seq).Selected = True
     txtFileName = lvFiles.ListItems(seq).Text
     Call mnuPreview_Click
     Exit Sub
   End Select
  Next seq
End Sub

Private Sub cmdPrevCancel_Click()
 pbxPrev.Visible = False
 icmDrives.Enabled = True
 Toolbar1.Enabled = True
 cmbFileTypes.Enabled = True
 cmdCancel.Enabled = True
 cmdOKay.Enabled = True
 lvFiles.SelectedItem.EnsureVisible
 lvFiles.SetFocus
End Sub




Private Sub Form_Load()
 Comm1.Width = 6345
 Comm1.Height = 5070
 Comm1.Top = (Screen.Height - Comm1.Height) / 2
 Comm1.Left = (Screen.Width - Comm1.Width) / 2
 
 With cmbFileTypes '***Change pattern of files to display in dropdown here***
  .List(0) = "All Files (*.*)"
  .AddItem "Flash Movies (*.swf)"
  .AddItem "Flash Files (*.fla)"
  .AddItem "Flash Movies & Files (*.swf;*.fla)"
  .AddItem "Image Files (*.bmp;*.jpg;*.wmf;*.gif;*.ico;*.cur)"
  .AddItem "Masks (*.msk)"
  .ListIndex = 0
 End With


'Text1.Text = "c:\_Osmax_3\CreatorV9\lighttab" ' ***Set initial path on loading here***
On Error GoTo NoPath
Dir1.Path = Text1.Text


Toolbar1.Buttons(1).Enabled = False

'Size the picture boxes containing the icons
pic16.Width = (SMALL_ICON) * Screen.TwipsPerPixelX
pic16.Height = (SMALL_ICON) * Screen.TwipsPerPixelY
pic32.Width = LARGE_ICON * Screen.TwipsPerPixelX
pic32.Height = LARGE_ICON * Screen.TwipsPerPixelY


If Text1.Text = "" Then
Text1.Text = Left(Drive1, 2)
 If Right(Text1.Text, 1) <> "\" Then Text1.Text = Text1.Text & "\"
Else
 If Right(Text1.Text, 1) <> "\" Then Text1.Text = Text1.Text & "\"
End If

Call DisplayDrives
Call Display

 If Len(Text1.Text) = 3 Then
  Toolbar1.Buttons(1).Enabled = False
 Else
  Toolbar1.Buttons(1).Enabled = True
 End If
Exit Sub

NoPath:
Text1.Text = ""
Resume Next
End Sub

Private Function GetIcon(FileName As String, Index As Long, ImageName As String) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
'On Error Resume Next
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
Dim r As Long

Label1.Caption = Index

'Get a handle to the small icon
hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'Get a handle to the large icon
hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

'If the handle(s) exists, load it into the picture box(es)
If hLIcon <> 0 Then
  'Large Icon
  With pic32
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hLIcon, ShInfo.iIcon, pic32.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  'Small Icon
  With pic16
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hSIcon, ShInfo.iIcon, pic16.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  Set imgObj = iml32.ListImages.Add(Index, ImageName, pic32.Image)
  Set imgObj = iml16.ListImages.Add(Index, ImageName, pic16.Image)
End If
End Function


Private Sub ShowIcons(LV As Object)
'-----------------------------------------
'Show the icons in the lvw
'-----------------------------------------
On Error Resume Next

Dim Item As ListItem
With LV
  '.ListItems.Clear
  .Icons = iml32        'Large
  .SmallIcons = iml16   'Small
  For Each Item In .ListItems
    If GetAttr(Item.Key) = vbDirectory Then
     Item.Icon = "folder"
     Item.SmallIcon = "folder"
    Else
     If UCase(Right(Item.Key, 4)) = ".EXE" Or UCase(Right(Item.Key, 4)) = ".ICO" Or UCase(Right(Item.Key, 4)) = ".CUR" Then
      Item.Icon = Item.Key
      Item.SmallIcon = Item.Key
     Else
      Item.Icon = UCase(Right(Item.Key, 4))
      Item.SmallIcon = UCase(Right(Item.Key, 4))
     End If
    End If
  Next
End With

End Sub
Private Sub GetAllIcons(LV As Object)
'--------------------------------------------------
'Extract all icons
'--------------------------------------------------
Dim FileName As String
Dim IconCount As Long
Dim GotFolder As Boolean

On Local Error Resume Next
Dim i As Long
List1.Clear
For i = 1 To LV.ListItems.Count
 CheckType = GetAttr(LV.ListItems(i).Key)
 If CheckType = vbDirectory Then
  If GotFolder = False Then
   IconCount = IconCount + 1
   GotFolder = True
   FileName = LV.ListItems(i).Key
   GetIcon FileName, IconCount, "folder"
  End If
 Else
  FileName = LV.ListItems(i).Key
  FileExt = UCase(Right(FileName, 4))
  If FileExt = ".EXE" Or FileExt = ".ICO" Or FileExt = ".CUR" Then
   IconCount = IconCount + 1
   GetIcon FileName, IconCount, FileName
  Else
   Dim ExtFound As Boolean
   For seq = 0 To List1.ListCount - 1
    If FileExt = List1.List(seq) Then ExtFound = True
   Next seq
   If ExtFound = False Then
    IconCount = IconCount + 1
    List1.AddItem FileExt
    GetIcon FileName, IconCount, FileExt
   Else
    ExtFound = False
   End If
  End If
 End If
Next i
Label2.Caption = IconCount
End Sub

Private Sub Form_Unload(Cancel As Integer)
 icmDrives.Enabled = True
 Toolbar1.Enabled = True
 cmbFileTypes.Enabled = True
 cmdCancel.Enabled = True
 cmdOKay.Enabled = True
End Sub

Private Sub icmDrives_Click()
 Dim TempDrive As String
 TempDrive = Text1.Text
 On Error GoTo NoDrive:
 Text1.Text = icmDrives.SelectedItem.Key
 If Right(Text1.Text, 1) <> "\" Then Text1.Text = Text1.Text & "\"
 
 If icmDrives.ComboItems.Count > icmDrives.SelectedItem.Index Then
  For seq = icmDrives.SelectedItem.Index + 1 To icmDrives.ComboItems.Count
   If Len(icmDrives.ComboItems(seq).Key) = 3 Then
    Exit For
   End If
  Next seq
  For seq2 = (seq - 1) To (icmDrives.SelectedItem.Index + 1) Step -1
   If Len(icmDrives.ComboItems(seq2).Key) <> 3 Then
    icmDrives.ComboItems.Remove seq2
   End If
  Next seq2
 End If
 
 Call Display
 If Len(Text1.Text) = 3 Then
  Toolbar1.Buttons(1).Enabled = False
 Else
  Toolbar1.Buttons(1).Enabled = True
 End If
 Exit Sub
 
NoDrive:
Msg = "Device Unavailable"
Style = vbokobly + vbCritical
Title = "Device Unavailable"
Response = MsgBox(Msg, Style, Title)
Text1.Text = TempDrive
icmDrives.ComboItems(Left(Text1.Text, 3)).Selected = True

End Sub


Private Sub lvFiles_DblClick()
 If lvFiles.ListItems.Count = 0 Then Exit Sub
' If File2.ListCount = 0 And Dir1.ListCount = 0 Then Exit Sub
 
 If Mid(lvFiles.SelectedItem.Key, Len(lvFiles.SelectedItem.Key) - 3, 1) <> "." Then
  Dir1.Path = lvFiles.SelectedItem.Key
  File2.Path = lvFiles.SelectedItem.Key
  Text1.Text = lvFiles.SelectedItem.Key
  Text2.Text = lvFiles.SelectedItem.Text
  If Right(Text1.Text, 1) <> "\" Then Text1.Text = Text1.Text & "\"
  lvFiles.ListItems.Clear
  If Len(Text1.Text) > 3 Then Toolbar1.Buttons(1).Enabled = True
  For seq = 1 To icmDrives.ComboItems.Count
   If icmDrives.ComboItems(seq).Selected = True Then
    icmDrives.ComboItems.Add seq + 1, Text1.Text, Text2.Text, "open folder", , icmDrives.ComboItems(seq).Indentation + 1
    icmDrives.ComboItems(seq).Selected = False
    icmDrives.ComboItems(seq + 1).Selected = True
    Exit For
   End If
  Next seq
  Call Display
 
 Else
  Dim i As Integer
  Dim TotalFiles As Integer
  For i = 1 To lvFiles.ListItems.Count
   If lvFiles.ListItems(i).Selected = True Then
      If Mid(lvFiles.SelectedItem.Key, Len(lvFiles.SelectedItem.Key) - 3, 1) = "." Then
       TotalFiles = TotalFiles + 1
      End If
   End If
  Next i
  If TotalFiles = 1 Then Call cmdOKay_Click
 End If
End Sub


Private Sub lvFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then
  For i = 1 To lvFiles.ListItems.Count
    lvFiles.ListItems(i).Selected = False
  Next i
 End If
End Sub

Private Sub lvFiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'On Error Resume Next
 'If File2.ListCount = 0 And Dir1.ListCount = 0 Then Exit Sub
 If lvFiles.ListItems.Count = 0 Then Exit Sub
 
 If Button = 1 Then
  txtFileName = ""
  If Mid(lvFiles.SelectedItem.Key, Len(lvFiles.SelectedItem.Key) - 3, 1) = "." Then
   For i = 1 To lvFiles.ListItems.Count
    If lvFiles.ListItems(i).Selected = True Then
     txtFileName = txtFileName & lvFiles.ListItems(i).Text & " "
    End If
   Next i
  End If
 End If
 
 If Button = 2 Then
  txtFileName = ""
  If Mid(lvFiles.SelectedItem.Key, Len(lvFiles.SelectedItem.Key) - 3, 1) = "." Then
     txtFileName = lvFiles.SelectedItem.Text
     Select Case UCase(Right(lvFiles.SelectedItem.Text, 4))
      Case ".BMP", ".JPG", ".ICO", ".GIF", ".CUR", ".WMF"
      mnuPreview.Enabled = True
      Case Else
      mnuPreview.Enabled = False
     End Select
     PopupMenu mnuOptions
  End If
  
 End If

 Exit Sub
 


End Sub


Private Sub mnuPreview_Click()
 On Error Resume Next
 imgPrev.Picture = LoadPicture("")
 imgHolder.Picture = LoadPicture("")
 icmDrives.Enabled = False
 Toolbar1.Enabled = False
 cmbFileTypes.Enabled = False
 cmdCancel.Enabled = False
 cmdOKay.Enabled = False
 
 Dim imgWidth As Integer
 Dim imgHeight As Integer
 Dim imgWFact As Single
 Dim imgHFact As Single
 Dim imgFactor As Single
 pbxPrev.Visible = True
 imgHolder = LoadPicture(lvFiles.SelectedItem.Key)
 Label7.Caption = FileDateTime(lvFiles.SelectedItem.Key)
 Label8.Caption = txtFileName
 Label9.Caption = Format(FileLen(lvFiles.SelectedItem.Key), "#,###") & " bytes"
 Label10.Caption = "(" & Format(FileLen(lvFiles.SelectedItem.Key) / 1000, "#,###") & " Kb)"
 imgWidth = imgHolder.Width
 imgHeight = imgHolder.Height
 imgWFact = ((pbxPrev.Width) - 2815) / imgWidth
 imgHFact = ((pbxPrev.Height) - 720) / imgHeight
 If imgWFact < imgHFact Then
  imgFactor = imgWFact
 Else
  imgFactor = imgHFact
 End If
 If imgFactor > 1 Then imgFactor = 1
 imgPrev.Width = imgWidth * imgFactor
 imgPrev.Height = imgHeight * imgFactor
 imgPrev.Left = ((pbxPrev.Width - imgPrev.Width) / 2) - 1260
 imgPrev.Top = ((pbxPrev.Height - imgPrev.Height) / 2) - 180
 imgPrev.Picture = imgHolder.Picture
 
 cmdNext.Enabled = False
 cmdPrev.Enabled = False
 
 If lvFiles.SelectedItem.Index < lvFiles.ListItems.Count Then
  For seq = lvFiles.SelectedItem.Index + 1 To lvFiles.ListItems.Count
   Select Case UCase(Right(lvFiles.ListItems(seq).Text, 4))
    Case ".BMP", ".JPG", ".ICO", ".GIF", ".CUR", ".WMF"
     cmdNext.Enabled = True
     Exit For
   End Select
  Next seq
 End If
 
 If lvFiles.SelectedItem.Index > 0 Then
  For seq = lvFiles.SelectedItem.Index - 1 To 1 Step -1
   Select Case UCase(Right(lvFiles.ListItems(seq).Text, 4))
    Case ".BMP", ".JPG", ".ICO", ".GIF", ".CUR", ".WMF"
     cmdPrev.Enabled = True
     Exit For
   End Select
  Next seq
 End If
 
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Index
  Case 1
   If Len(icmDrives.SelectedItem.Key) = 3 Then Exit Sub
   Dim RemItem As Integer
   RemItem = icmDrives.SelectedItem.Index
   icmDrives.ComboItems(RemItem).Selected = False
   icmDrives.ComboItems(RemItem - 1).Selected = True
   icmDrives.ComboItems.Remove RemItem
   Call icmDrives_Click
   'Dim CheckOK As String
   'CheckOK = Left(Text1.Text, Len(Text1.Text) - 1)
   'If CheckOK = Left(Drive1, 2) Then Exit Sub
   'For seq = Len(CheckOK) To 1 Step -1
   'If Mid(CheckOK, seq, 1) = "\" Then
   ' Text1.Text = Left(CheckOK, seq)
   ' Exit For
   'End If
   'Next seq
   'If Len(Text1.Text) = 3 Then Toolbar1.Buttons(1).Enabled = False
   'Call Display
  Case 2
   On Error GoTo DupFile
   Dim dctName As String
   dctName = InputBox("Please enter a name for the new folder", "Create New Folder", "")
   If (dctName = "") Then
    Exit Sub
   Else
    MkDir Text1.Text & dctName
   End If
   Dir1.Refresh
   Call Display
   For seq = 1 To lvFiles.ListItems.Count
    lvFiles.ListItems(seq).Selected = False
   Next seq
   Set lvFiles.SelectedItem = lvFiles.ListItems(Text1.Text & dctName)
   lvFiles.SelectedItem.EnsureVisible
   lvFiles.SetFocus
  Case 3
   If lvFiles.View = 0 Then
    lvFiles.View = 2
    Toolbar1.Buttons(3).ButtonMenus(1).Enabled = False
    Toolbar1.Buttons(3).ButtonMenus(2).Enabled = True
    Toolbar1.Buttons(3).Image = 4
    Toolbar1.Buttons(3).ToolTipText = "Large Icons"
   ElseIf lvFiles.View = 2 Then
    lvFiles.View = 0
    Toolbar1.Buttons(3).ButtonMenus(1).Enabled = True
    Toolbar1.Buttons(3).ButtonMenus(2).Enabled = False
    Toolbar1.Buttons(3).Image = 3
    Toolbar1.Buttons(3).ToolTipText = "Small Icons"
   End If
 End Select
 Exit Sub
 
DupFile:
 MsgBox "The new folder name is invalid", vbOKOnly + vbCritical, "Invalid folder name"
 Exit Sub
End Sub


Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
 Select Case ButtonMenu.Index
  Case 1
   If lvFiles.View <> 2 Then lvFiles.View = 2
   Toolbar1.Buttons(3).ButtonMenus(1).Enabled = False
   Toolbar1.Buttons(3).ButtonMenus(2).Enabled = True
   Toolbar1.Buttons(3).Image = 4
   Toolbar1.Buttons(3).ToolTipText = "Large Icons"
  Case 2
   If lvFiles.View <> 0 Then lvFiles.View = 0
   Toolbar1.Buttons(3).ButtonMenus(2).Enabled = False
   Toolbar1.Buttons(3).ButtonMenus(1).Enabled = True
   Toolbar1.Buttons(3).Image = 3
   Toolbar1.Buttons(3).ToolTipText = "Small Icons"
 End Select
End Sub


