VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmmain 
   BackColor       =   &H8000000C&
   Caption         =   "VB File Compare v1.0  <DarkMercenary44@cheapconnect.net>"
   ClientHeight    =   7200
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9780
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   0
      ScaleHeight     =   6255
      ScaleWidth      =   1980
      TabIndex        =   8
      Top             =   945
      Width           =   1980
      Begin RichTextLib.RichTextBox mstats 
         Height          =   2475
         Left            =   45
         TabIndex        =   9
         Top             =   210
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   4366
         _Version        =   393217
         Enabled         =   0   'False
         ReadOnly        =   -1  'True
         TextRTF         =   $"frmmain.frx":0E42
      End
      Begin RichTextLib.RichTextBox cstats 
         Height          =   2475
         Left            =   45
         TabIndex        =   11
         Top             =   2970
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   4366
         _Version        =   393217
         Enabled         =   0   'False
         ReadOnly        =   -1  'True
         TextRTF         =   $"frmmain.frx":0EC4
      End
      Begin VB.Label Label2 
         Caption         =   "Compare File Stats"
         Enabled         =   0   'False
         Height          =   210
         Left            =   300
         TabIndex        =   12
         Top             =   2760
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Master File Stats"
         Enabled         =   0   'False
         Height          =   210
         Left            =   360
         TabIndex        =   10
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox pbTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   9780
      TabIndex        =   0
      Top             =   0
      Width           =   9780
      Begin VB.CommandButton cmdhelp 
         Caption         =   "Help"
         Height          =   285
         Left            =   5025
         TabIndex        =   13
         Top             =   600
         Width           =   1245
      End
      Begin VB.CommandButton htile 
         Caption         =   "Horizontal Tiling"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6870
         TabIndex        =   7
         Top             =   330
         Width           =   2385
      End
      Begin VB.CommandButton vtile 
         Caption         =   "Vertical Tiling"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6870
         TabIndex        =   6
         Top             =   60
         Width           =   2385
      End
      Begin VB.CommandButton begin 
         Caption         =   "Begin Comparision"
         Height          =   210
         Left            =   30
         TabIndex        =   5
         Top             =   675
         Width           =   4935
      End
      Begin VB.CommandButton browse2 
         Caption         =   "Browse"
         Height          =   285
         Left            =   5025
         TabIndex        =   4
         Top             =   330
         Width           =   1245
      End
      Begin VB.CommandButton browse1 
         Caption         =   "Browse"
         Height          =   285
         Left            =   5025
         TabIndex        =   3
         Top             =   60
         Width           =   1245
      End
      Begin VB.TextBox txtcompare 
         Height          =   285
         Left            =   30
         TabIndex        =   2
         Text            =   "Compare File"
         Top             =   330
         Width           =   4935
      End
      Begin VB.TextBox txtmaster 
         Height          =   285
         Left            =   30
         TabIndex        =   1
         Text            =   "Master File"
         Top             =   60
         Width           =   4935
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   9225
      Top             =   1020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit








Private Sub begin_Click()

If txtmaster = "Master File" Or txtcompare = "Compare File" Then
    MsgBox "You must specify a Master and Compare File", vbCritical, "YOU SCREWED UP"
    Exit Sub
End If

Dim masterline As String
Dim compareline As String
Dim mlinecount As Integer
Dim clinecount As Integer
Dim diffcount As Integer
Dim I As Integer
Dim K As Integer

newdoc(0).Show
newdoc(0).Caption = "Master  <" & txtmaster.Text & ">"
newdoc(1).Show
newdoc(1).Caption = "Compare  <" & txtcompare.Text & ">"
vtile.Enabled = True
htile.Enabled = True
mtxtadd ("===========================================================")
mtxtadd ("Master File: " & txtmaster.Text)
mtxtadd ("===========================================================")
mtxtadd (" ")
mtxtadd (" ")
ctxtadd ("===========================================================")
ctxtadd ("Compare File: " & txtcompare.Text)
ctxtadd ("===========================================================")
ctxtadd (" ")
ctxtadd (" ")



Open txtmaster.Text For Input As #1
Do Until EOF(1)
    Line Input #1, masterline
    mlinecount = mlinecount + 1
Loop
Close #1

ReDim masterarray(mlinecount)
mlinecount = 0

Open txtmaster.Text For Input As #1
Do Until EOF(1)
    Line Input #1, masterline
    masterarray(mlinecount) = masterline
    mtxtadd (masterarray(mlinecount))
    mlinecount = mlinecount + 1
Loop
Close #1
    
Open txtcompare.Text For Input As #1
Do Until EOF(1)
    Line Input #1, compareline
    clinecount = clinecount + 1
Loop
Close #1

ReDim comparearray(clinecount)
ReDim diffarray(clinecount)
clinecount = 0

Open txtcompare.Text For Input As #1
Do Until EOF(1)
    Line Input #1, compareline
    comparearray(clinecount) = compareline
    ctxtadd (comparearray(clinecount))
    clinecount = clinecount + 1
Loop
Close #1
comparecount = clinecount

mastertext = newdoc(0).rtext.Text
comparetext = newdoc(1).rtext.Text

DoEvents

Call mstatsadd("Lines", mlinecount)
Call cstatsadd("Lines", clinecount)

For I = 0 To clinecount
    If InStr(1, mastertext, comparearray(I)) = 0 Then
        With newdoc(1).rtext
            .SelStart = InStr(1, comparetext, comparearray(I)) - 1
            .SelLength = Len(comparearray(I))
            .SelColor = vbRed
        End With
        diffcount = diffcount + 1
        diffarray(diffcount) = comparearray(I)
    End If
Next I
DoEvents
Call cstatsadd("Differences", diffcount)
End Sub


Private Sub browse1_Click()
cd1.Filter = "Modules (*.bas)|*.bas|Forms (*.frm)|*.frm"
cd1.ShowOpen
txtmaster.Text = cd1.FileName
End Sub

Private Sub browse2_Click()
cd1.Filter = "Modules (*.bas)|*.bas|Forms (*.frm)|*.frm"
cd1.ShowOpen
txtcompare.Text = cd1.FileName
End Sub


Private Sub cmdhelp_Click()
Form1.Show vbModal

End Sub

Private Sub htile_Click()
Me.Arrange vbTileHorizontal
End Sub

Private Sub vtile_Click()
Me.Arrange vbTileVertical
End Sub


