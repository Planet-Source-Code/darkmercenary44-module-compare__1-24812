VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmbuildwordlist 
   Caption         =   "Build Word List"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   Icon            =   "frmbuildwordlist.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstGo 
      Height          =   1860
      Left            =   75
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   4695
      Width           =   6555
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   315
      Left            =   5085
      TabIndex        =   5
      Top             =   465
      Width           =   1515
   End
   Begin VB.ListBox lstShow 
      Height          =   1860
      ItemData        =   "frmbuildwordlist.frx":0E42
      Left            =   75
      List            =   "frmbuildwordlist.frx":0E44
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   2445
      Width           =   6555
   End
   Begin VB.Frame Frame1 
      Height          =   1350
      Left            =   75
      TabIndex        =   3
      Top             =   810
      Width           =   6555
      Begin MSComctlLib.ProgressBar cpb 
         Height          =   225
         Left            =   255
         TabIndex        =   13
         Top             =   405
         Visible         =   0   'False
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.CommandButton cmdBuild 
         Caption         =   "Build List"
         Height          =   315
         Left            =   4950
         TabIndex        =   6
         Top             =   945
         Visible         =   0   'False
         Width           =   1515
      End
      Begin MSComctlLib.ProgressBar tpb 
         Height          =   225
         Left            =   255
         TabIndex        =   15
         Top             =   870
         Visible         =   0   'False
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label6 
         Caption         =   "Total Progress"
         Height          =   195
         Left            =   1965
         TabIndex        =   16
         Top             =   660
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label5 
         Caption         =   "Current File Progress"
         Height          =   225
         Left            =   1785
         TabIndex        =   14
         Top             =   180
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label lblaftersearch 
         Caption         =   $"frmbuildwordlist.frx":0E46
         Height          =   465
         Left            =   5430
         TabIndex        =   12
         Top             =   330
         Visible         =   0   'False
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdStartSearch 
      Caption         =   "Start Search"
      Height          =   315
      Left            =   5085
      TabIndex        =   2
      Top             =   90
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   315
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "*.txt"
      Top             =   465
      Width           =   4815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   4815
   End
   Begin VB.Label Label4 
      Caption         =   "0 file(s) in list"
      Height          =   225
      Left            =   75
      TabIndex        =   11
      Top             =   4470
      Visible         =   0   'False
      Width           =   3510
   End
   Begin VB.Label Label3 
      Caption         =   "lstShow"
      Height          =   225
      Left            =   75
      TabIndex        =   10
      Top             =   2190
      Visible         =   0   'False
      Width           =   3510
   End
   Begin VB.Label Label2 
      Caption         =   "Searching...."
      Height          =   210
      Left            =   4905
      TabIndex        =   8
      Top             =   6750
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label Label1 
      Height          =   210
      Left            =   75
      TabIndex        =   7
      Top             =   6750
      Width           =   4005
   End
   Begin VB.Menu mnlstShow 
      Caption         =   "lstShow"
      Visible         =   0   'False
      Begin VB.Menu mnShowAddcheckedfilestofinallist 
         Caption         =   "Add checked files to final list"
      End
      Begin VB.Menu mnShowSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnShowSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnShowDeselectall 
         Caption         =   "Deselect All"
      End
      Begin VB.Menu mnShowSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnShowCancel 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnlstGo 
      Caption         =   "lstGo"
      Visible         =   0   'False
      Begin VB.Menu mnGoRemovecheckedfilesfromlist 
         Caption         =   "Remove checked files from final list"
      End
      Begin VB.Menu mnGoSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnGoSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnGoDeselectAll 
         Caption         =   "Deselect All"
      End
      Begin VB.Menu mnGoSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnGoCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmbuildwordlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CFiles As New colFiles
Dim lstGoCount As Integer
Dim lstShowCount As Integer
Private Sub startsearch_Click()

End Sub


Private Sub cmdBuild_Click()
Dim temptext As String
Dim linecount As Integer
Dim I As Integer

lblaftersearch.Visible = False
lblaftersearch.Refresh
Frame1.Refresh
Label5.Visible = True
Label6.Visible = True
cpb.Visible = True
tpb.Visible = True
tpb.Value = 0
cpb.Value = 0
Frame1.Refresh
tpb.Max = lstGo.ListCount
'tpb.Value = 1
For I = 0 To lstGo.ListCount - 1
    Open lstGo.List(I) For Input As #1
        Do Until EOF(1)
            Line Input #1, temptext
            linecount = linecount + 1
        Loop
    Close #1
    cpb.Max = linecount
    Open lstGo.List(I) For Input As #1
        Do Until EOF(1)
            Line Input #1, temptext
            Call parsedata(temptext, "wordlist.txt")
            cpb.Value = cpb.Value + 1
        Loop
        tpb.Value = tpb.Value + 1
        cpb.Value = 0
        linecount = 0
    Close #1
Next I
cpb.Value = cpb.Max
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdStartSearch_Click()

Dim data As String

Label2.Visible = True
Label2.Refresh
data = Drive1.Drive
cmdStartSearch.Enabled = False
cmdExit.Enabled = False
Me.MousePointer = 11
If InStr(data, " ") Then
    data = Left(data, InStr(data, " ") - 1)
End If

CFiles.Clear
CFiles.LoadFiles data & "\" & Text1.Text, True

    
Dim l As Long
For l = 1 To CFiles.Count
    lstShow.AddItem CFiles(l).sPath & CFiles(l).sNameAndExtension
Next l

Label1.Caption = "Found " & CFiles.Count & " file(s)"

'Label1.Caption = FindFile("*.txt", data, Me, lstShow)
Label2.Caption = "Search Complete"
Label3.Visible = True
Label4.Visible = True
lstShowCount = CFiles.Count
Label3.Caption = lstShowCount & " file(s) in list"
Label4.Caption = lstGoCount & " file(s) in list"
With lblaftersearch
    .Visible = True
    .Height = 870
    .Left = 75
    .Top = 180
    .Width = 4758
End With
cmdStartSearch.Enabled = True
cmdExit.Enabled = True
Me.MousePointer = 0
End Sub


Private Sub lstGo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu mnlstGo
End Sub


Private Sub lstShow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu mnlstShow
End Sub


Private Sub mnGoDeselectAll_Click()
Dim I As Integer

For I = 0 To lstGo.ListCount - 1
    lstGo.Selected(I) = False
Next I
End Sub

Private Sub mnGoRemovecheckedfilesfromlist_Click()
Dim I As Integer
Dim K As Integer

For K = 1 To lstGo.SelCount
    For I = 0 To lstGo.ListCount - 1
        If lstGo.Selected(I) = True Then
            lstShow.AddItem lstGo.List(I)
            lstShowCount = lstShowCount + 1
            lstGo.RemoveItem I
            lstGoCount = lstGoCount - 1
            Exit For
        End If
    Next I
Next K


Label3.Caption = lstShowCount & " file(s) in list"

Label4.Caption = lstGoCount & " file(s) in list"
End Sub

Private Sub mnGoSelectAll_Click()
Dim I As Integer

For I = 0 To lstGo.ListCount - 1
    lstGo.Selected(I) = True
Next I
End Sub


Private Sub mnShowAddcheckedfilestofinallist_Click()

Dim I As Integer
Dim K As Integer

For K = 1 To lstShow.SelCount
    For I = 0 To lstShow.ListCount - 1
        If lstShow.Selected(I) = True Then
            lstGo.AddItem lstShow.List(I)
            lstGoCount = lstGoCount + 1
            lstShow.RemoveItem I
            lstShowCount = lstShowCount - 1
            Exit For
        End If
    Next I
Next K


Label3.Caption = lstShowCount & " file(s) in list"
cmdBuild.Visible = True
Label4.Caption = lstGoCount & " file(s) in list"

End Sub

Private Sub mnShowDeselectall_Click()
Dim I As Integer

For I = 0 To lstShow.ListCount - 1
    lstShow.Selected(I) = False
Next I
End Sub

Private Sub mnShowSelectAll_Click()
Dim I As Integer

For I = 0 To lstShow.ListCount - 1
    lstShow.Selected(I) = True
Next I
End Sub


