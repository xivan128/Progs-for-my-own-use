Sub taskcpy()
Dim myString As String


sepStr = ","""
interstring = ",,,,,,,,"
tail = "4,1,ivan (15417572),,,en,Europe/Moscow"
For m = 2 To 128
    s1 = Workbooks("test.csv").Sheets("test2").Cells(m, "A")
    k = m * 2 - 1
    n = InStr(1, s1, sepStr, vbTextCompare)
    Workbooks("test.csv").Sheets("test").Cells(k - 1, "A") = "task," + Left(s1, n) + tail
    Sheets("test").Cells(k, "A") = interstring
Next m
End Sub
