Attribute VB_Name = "Module1"

Option Explicit

Public newdoc(1) As New frmdoc
Public masterarray() As String
Public comparearray() As String
Public diffarray() As String
Public mastertext As String
Public comparetext As String
Public comparecount As Integer

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Function Wait(ByVal TimeToWait As Long) 'Time In seconds
    Dim EndTime As Long
    EndTime = GetTickCount + TimeToWait * 1000 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds


    Do Until GetTickCount > EndTime


        DoEvents
        Loop
    End Function

Public Sub mtxtadd(data As String)
newdoc(0).rtext.Text = newdoc(0).rtext.Text & data & vbCrLf
End Sub

Public Sub ctxtadd(data As String)
newdoc(1).rtext.Text = newdoc(1).rtext.Text & data & vbCrLf
End Sub

Public Sub mstatsadd(heading As String, data As Variant)

Dim headingdata As String
Dim datadata As Variant

headingdata = heading & ": "
datadata = data & vbCr


If frmmain.mstats.Enabled = False Then frmmain.mstats.Enabled = True

With frmmain.mstats
    .Locked = False
    .SetFocus
    SendKeys headingdata & datadata
    DoEvents
    .SelStart = Len(.Text) - Len(headingdata & datadata) - 1
    .SelLength = Len(headingdata)
    .SelColor = vbRed
    .SetFocus
    SendKeys "^{END}"
    .Locked = True
End With
End Sub

Public Sub cstatsadd(heading As String, data As Variant)
Dim headingdata As String
Dim datadata As Variant

headingdata = heading & ": "
datadata = data & vbCr


If frmmain.cstats.Enabled = False Then frmmain.cstats.Enabled = True

With frmmain.cstats
    .Locked = False
    .SetFocus
    SendKeys headingdata & datadata
    DoEvents
    .SelStart = Len(.Text) - Len(headingdata & datadata) - 1
    .SelLength = Len(headingdata)
    .SelColor = vbRed
    .SetFocus
    SendKeys "^{END}"
    .Locked = True
End With
End Sub
