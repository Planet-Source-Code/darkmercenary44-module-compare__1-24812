VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmdoc 
   Caption         =   "Untitled..."
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3780
   Icon            =   "frmdoc.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2790
   ScaleWidth      =   3780
   Begin RichTextLib.RichTextBox rtext 
      Height          =   2805
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   4948
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"frmdoc.frx":058A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnedit 
      Caption         =   "edit"
      Visible         =   0   'False
      Begin VB.Menu mnJump 
         Caption         =   "Jump to line in Master File"
      End
      Begin VB.Menu mnJumpCompare 
         Caption         =   "Jump to line in Compare File"
      End
   End
End
Attribute VB_Name = "frmdoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnFile1_Click()

End Sub


Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
With rtext
    .Top = 0 + NewHeight
    .Left = 0
    .Width = Me.Width - 125
    .Height = Me.Height - 400 - NewHeight
End With
End Sub

Private Sub Form_Load()

With rtext
    .Top = 0
    .Left = 0
    .Width = Me.Width - 125
    .Height = Me.Height - 400
End With

End Sub


Private Sub Form_Resize()
If Me.WindowState <> 1 Then
    With rtext
        .Top = 0
        .Left = 0
        .Width = Me.Width - 125
        .Height = Me.Height - 400
    End With
   
End If
End Sub


Private Sub mnJump_Click()
Dim compareline As String
Dim I As Integer
compareline = rtext.SelText

If InStr(1, mastertext, compareline) = 0 Then
    I = 0
Else
    I = InStr(1, mastertext, compareline)
End If

With newdoc(0).rtext
    .SetFocus
    .SelStart = I - 1
    .SelLength = Len(compareline)
End With
End Sub

Private Sub mnJumpCompare_Click()
Dim masterline As String
Dim I As Integer

masterline = rtext.SelText

If InStr(1, comparetext, masterline) = 0 Then
    I = 0
Else
    I = InStr(1, comparetext, masterline)
End If

With newdoc(1).rtext
    .SetFocus
    .SelStart = I - 1
    .SelLength = Len(masterline)
End With
End Sub


Private Sub rtext_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    SendKeys "{HOME}"
    SendKeys "+{END}"
End If
If Button = 2 Then
    PopupMenu mnedit
End If
End Sub


