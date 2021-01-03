VERSION 5.00
Begin VB.Form frmFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FileSystemObject Sample"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   11
      Left            =   6240
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   6645
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   10
      Left            =   6240
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   6285
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   9
      Left            =   6240
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   5925
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   8
      Left            =   6240
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   5565
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   7
      Left            =   6240
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   5205
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   6
      Left            =   6240
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   4845
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   5
      Left            =   1680
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   6645
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   4
      Left            =   1680
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   6285
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   5925
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   5565
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   5205
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   4845
      Width           =   2655
   End
   Begin VB.ListBox lstFiles 
      Height          =   4350
      Left            =   5430
      TabIndex        =   2
      Top             =   360
      Width           =   4485
   End
   Begin VB.ListBox lstFolders 
      Height          =   4350
      Left            =   2445
      TabIndex        =   1
      Top             =   360
      Width           =   2955
   End
   Begin VB.ListBox lstDrives 
      Height          =   4350
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lbl 
      Caption         =   "Files:"
      Height          =   255
      Index           =   14
      Left            =   5520
      TabIndex        =   29
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lbl 
      Caption         =   "Folders:"
      Height          =   255
      Index           =   13
      Left            =   2520
      TabIndex        =   28
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lbl 
      Caption         =   "Drives:"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   255
      Index           =   11
      Left            =   4680
      TabIndex        =   14
      Top             =   6720
      Width           =   1500
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   255
      Index           =   10
      Left            =   4680
      TabIndex        =   13
      Top             =   6360
      Width           =   1500
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   255
      Index           =   9
      Left            =   4680
      TabIndex        =   12
      Top             =   6000
      Width           =   1500
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   255
      Index           =   8
      Left            =   4680
      TabIndex        =   11
      Top             =   5640
      Width           =   1500
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   255
      Index           =   7
      Left            =   4680
      TabIndex        =   10
      Top             =   5280
      Width           =   1500
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   255
      Index           =   6
      Left            =   4680
      TabIndex        =   9
      Top             =   4920
      Width           =   1500
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   6720
      Width           =   1500
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   6360
      Width           =   1500
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   1500
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   5640
      Width           =   1500
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   1500
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   1500
   End
End
Attribute VB_Name = "frmFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Module-level filesystem object:
Dim m_FileSys As FileSystemObject

'Show drives and folders
Private Sub Form_Load()
  
  Set m_FileSys = New FileSystemObject
  Dim Drive As Drive
  For Each Drive In m_FileSys.Drives
    lstDrives.AddItem Drive.Path
  Next
  
  lstDrives.ListIndex = 1
  
  Set Drive = m_FileSys.Drives(lstDrives.Text)
  
  FillFolderList Drive
  
End Sub

'Show Folders that are children of the root folder in
'the drive identified by DriveName parameter
Private Sub FillFolderList(ByVal Drive As Drive)

  lstFolders.Clear
  
  Dim Root As Folder
  Set Root = Drive.RootFolder
  
  Dim Folder As Folder
  For Each Folder In Root.SubFolders
    lstFolders.AddItem Folder.Name
  Next
  
End Sub

'Fill the file listbox with the names of files
'in the Folder parameter
Private Sub FillFileList(ByVal Folder As Folder)
  
  Dim File As File
  lstFiles.Clear
  
  For Each File In Folder.Files
    lstFiles.AddItem File.Name
  Next
  
End Sub

'Display details about a folder
Private Sub ShowFolderDetails(ByVal Folder As Folder)

  lbl(0) = "Attributes:"
  txt(0) = GetFileAttributeString(Folder.Attributes)
  lbl(1) = "Created:"
  txt(1) = Folder.DateCreated
  lbl(2) = "Last Accessed:"
  txt(2) = Folder.DateLastAccessed
  lbl(3) = "Last Modified:"
  txt(3) = Folder.DateLastModified
  lbl(4) = "Root Folder ?:"
  txt(4) = Folder.IsRootFolder
  lbl(5) = "Name:"
  txt(5) = Folder.Name
  lbl(6) = "Parent Folder:"
  txt(6) = Folder.ParentFolder
  lbl(7) = "Path:"
  txt(7) = Folder.Path
  lbl(8) = "Short Name:"
  txt(8) = Folder.ShortName
  lbl(9) = "Short Path:"
  txt(9) = Folder.ShortPath
  lbl(10) = "Size:"
  txt(10) = Folder.Size
  lbl(11) = "Type:"
  txt(11) = Folder.Type
  
End Sub

'Display details about a drive
Private Sub ShowDriveDetails(ByVal Drive As Drive)
  
  lbl(0) = "Available Space:"
  txt(0) = Drive.AvailableSpace
  lbl(1) = "Drive Letter:"
  txt(1) = Drive.DriveLetter
  lbl(2) = "Drive Type:"
  txt(2) = GetDriveTypeString(Drive.DriveType)
  lbl(3) = "Filesystem:"
  txt(3) = Drive.FileSystem
  lbl(4) = "Free Space:"
  txt(4) = Drive.FreeSpace
  lbl(5) = "Ready:"
  txt(5) = Drive.IsReady
  lbl(6) = "Path:"
  txt(6) = Drive.Path
  lbl(7) = "Rootfolder:"
  txt(7) = Drive.RootFolder.Name
  lbl(8) = "Serialnumber:"
  txt(8) = Drive.SerialNumber
  lbl(9) = "Share Name:"
  txt(9) = Drive.ShareName
  lbl(10) = "Total Size:"
  txt(10) = Drive.TotalSize
  lbl(11) = "Volume Name:"
  txt(11) = Drive.VolumeName
  
End Sub

'Show details about a file
Private Sub ShowFileDetails(ByVal File As File)

  lbl(0) = "Attributes:"
  txt(0) = GetFileAttributeString(File.Attributes)
  lbl(1) = "Created:"
  txt(1) = File.DateCreated
  lbl(2) = "Last Accessed:"
  txt(2) = File.DateLastAccessed
  lbl(3) = "Last Modified:"
  txt(3) = File.DateLastModified
  lbl(4) = "Drive:"
  txt(4) = File.Drive
  lbl(5) = "Name:"
  txt(5) = File.Name
  lbl(6) = "Parent Folder:"
  txt(6) = File.ParentFolder
  lbl(7) = "Path:"
  txt(7) = File.Path
  lbl(8) = "Short Name:"
  txt(8) = File.ShortName
  lbl(9) = "Short Path:"
  txt(9) = File.ShortPath
  lbl(10) = "Size:"
  txt(10) = File.Size
  lbl(11) = "Type:"
  txt(11) = File.Type
  
End Sub

'Show child folders and drive details
'when user clicks on a drive
Private Sub lstDrives_Click()
On Error GoTo ERR_ROUTINE

  Dim Drive As Drive
  Set Drive = m_FileSys.Drives(lstDrives.Text)
  
  If Not Drive.IsReady Then
    MsgBox "Drive not ready !", vbOKOnly Or vbCritical
    Exit Sub
  End If

  FillFolderList Drive
  ShowDriveDetails Drive
  
Exit Sub
ERR_ROUTINE:
  MsgBox Err.Description
  Resume Next
End Sub

'Show file details when user clicks on a particular file
Private Sub lstFiles_Click()
On Error GoTo ERR_ROUTINE

  Dim Drive As Drive
  Set Drive = m_FileSys.Drives(lstDrives.Text)
  
  If Not Drive.IsReady Then
    MsgBox "Drive not ready !", vbOKOnly Or vbCritical
    Exit Sub
  End If

  Dim Root As Folder
  Set Root = Drive.RootFolder
  
  Dim Folder As Folder
  Set Folder = Root.SubFolders(lstFolders.Text)

  Dim File As File
  Set File = Folder.Files(lstFiles.Text)
  
  ShowFileDetails File
  
Exit Sub
ERR_ROUTINE:
  MsgBox Err.Description
  Resume Next
End Sub

'Show folder details and fill
'file listbox with the names of files in the folder clicked
Private Sub lstFolders_Click()
On Error GoTo ERR_ROUTINE

  Dim Drive As Drive
  Set Drive = m_FileSys.Drives(lstDrives.Text)
  
  If Not Drive.IsReady Then
    MsgBox "Drive not ready !", vbOKOnly Or vbCritical
    Exit Sub
  End If

  Dim Root As Folder
  Set Root = Drive.RootFolder
  
  Dim Folder As Folder
  Set Folder = Root.SubFolders(lstFolders.Text)
  ShowFolderDetails Folder
  FillFileList Folder
  
Exit Sub
ERR_ROUTINE:
  MsgBox Err.Description
  Resume Next
End Sub

'Convert a DriveTypeConst to a string expression
'for display purposes
Private Function GetDriveTypeString(ByVal dType As DriveTypeConst) As String
  
  Select Case dType
    Case CDRom
      GetDriveTypeString = "CDRom"
    Case Fixed
      GetDriveTypeString = "Fixed Disk"
    Case RamDisk
      GetDriveTypeString = "Ram Disk"
    Case Remote
      GetDriveTypeString = "Remote Disk"
    Case Removable
      GetDriveTypeString = "Removable Disk"
    Case Else
      GetDriveTypeString = "Unknown Type"
  End Select
  
End Function

'Convert Attribute flag field to
'string expression for display purposes
Private Function GetFileAttributeString(ByVal Attr As FileAttribute) As String
  
  Dim s As String
  
  If Attr And Alias Then s = s & "Alias "
  If Attr And Archive Then s = s & "Archive "
  If Attr And Compressed Then s = s & "Compressed "
  If Attr And Directory Then s = s & "Directory "
  If Attr And Hidden Then s = s & "Hidden "
  If Attr And Normal Then s = s & "Normal "
  If Attr And ReadOnly Then s = s & "ReadOnly "
  If Attr And System Then s = s & "System "
  If Attr And Volume Then s = s & "Volume "
  GetFileAttributeString = s
  
End Function
