Attribute VB_Name = "Module1"
Public Sub Main()
Dim args As String
args = UCase$(Trim$(Command$))
Select Case Mid$(args, 1, 2)
Case "/C"
frmSet.Show

Case "/S"
frmMain.Show

End Select
End Sub

