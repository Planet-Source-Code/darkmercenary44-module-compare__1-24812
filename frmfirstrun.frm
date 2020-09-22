VERSION 5.00
Begin VB.Form frmfirstrun 
   Caption         =   "Scrambler First Run Setup..."
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   Icon            =   "frmfirstrun.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStartWordListBuilder 
      Caption         =   "Start Word List Builder"
      Height          =   315
      Left            =   4260
      TabIndex        =   7
      Top             =   2475
      Width           =   4470
   End
   Begin VB.CommandButton cmdCreateKey 
      Caption         =   "Create Key"
      Height          =   300
      Left            =   7350
      TabIndex        =   5
      Top             =   945
      Width           =   1395
   End
   Begin VB.TextBox txtKey 
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   4305
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1305
      Width           =   2985
   End
   Begin VB.TextBox txtKeyName 
      Height          =   285
      Left            =   4305
      TabIndex        =   2
      Top             =   960
      Width           =   2985
   End
   Begin VB.Label Label4 
      Caption         =   $"frmfirstrun.frx":0E42
      Height          =   645
      Left            =   165
      TabIndex        =   6
      Top             =   1785
      Width           =   8655
   End
   Begin VB.Label Label3 
      Caption         =   "This is your digital signature, Scrambler will store it."
      Height          =   240
      Left            =   645
      TabIndex        =   4
      Top             =   1350
      Width           =   3570
   End
   Begin VB.Label Label2 
      Caption         =   "Your Name as you want it to appear in digital signature."
      Height          =   225
      Left            =   315
      TabIndex        =   1
      Top             =   990
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   $"frmfirstrun.frx":0FAE
      Height          =   840
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   8790
   End
End
Attribute VB_Name = "frmfirstrun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCreateKey_Click()
txtKey.Text = GenKey(txtKeyName.Text)
With tbsettings
    .MoveFirst
    .Edit
    !Name = txtKeyName.Text
    !digitalkey = txtKey.Text
    .Update
End With
End Sub


Private Sub cmdStartWordListBuilder_Click()
Unload Me
frmbuildwordlist.Show
End Sub


