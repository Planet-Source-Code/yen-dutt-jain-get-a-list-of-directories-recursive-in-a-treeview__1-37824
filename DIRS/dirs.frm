VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List Folder v1.0"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   5160
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   9340
      _Version        =   393217
      Indentation     =   0
      Style           =   7
      Appearance      =   1
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   4800
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "List of Files in Selected Folder"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "List of Folders and Sub-Folders"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A(1000) As String
Public J As Integer
Public Counter As Integer
Public FolderName As String

Function ListDirectories(Path)
Dir1.Path = Path
For i = 0 To Dir1.ListCount - 1
   Dir1.ListIndex = i
   Categories = Dir1.List(i)
   If Len(Dir(Categories, vbDirectory)) > 1 Then ' Hidden, Read-Only, System files ommited
      A(J) = Categories
      TreeView1.Nodes.Add "x" & Counter, tvwChild, "x" & J, Dir(A(J), vbDirectory)
      J = J + 1
   End If
Next i
Counter = Counter + 1
If A(Counter) = "" Then
   'MsgBox ("Over")
Else
   ListDirectories (A(Counter))
End If
End Function

Private Sub File1_Click()
MsgBox (File1.Path & "\" & File1.FileName)
End Sub

Private Sub Form_Load()
FolderName = "C:\WINDOWS" 'It Must be a foldername not the root directory
J = 1
Counter = 0
TreeView1.Nodes.Add , tvwChild, "x" & Counter, Right(FolderName, Len(FolderName) - 3)
TreeView1.Nodes.Item(1).Expanded = True
ListDirectories (FolderName & "\")
File1.Path = Mid(FolderName, 1, 3)
End Sub

Private Sub TreeView1_Click()
File1.Path = Mid(FolderName, 1, 3) & TreeView1.SelectedItem.FullPath
End Sub
