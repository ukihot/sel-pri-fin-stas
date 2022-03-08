Attribute VB_Name = "Module1"

Public Function SheetDetect(SName As String) As Boolean
    Dim sht As Worksheet
    For Each sht In ThisWorkbook.Worksheets
        If sht.Name = SName Then
            SheetDetect = True
            Exit Function
        End If
    Next
End Function
