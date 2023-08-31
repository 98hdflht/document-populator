Attribute VB_Name = "Module1"
Sub studentInfo(wcn As Integer)
    lastrow = Worksheets("ELT Student Info").Cells(Rows.Count, "a").End(xlUp).Row
    Set searchrng = Worksheets("ELT Student Info").Range("a2:a" & lastrow).Find(wcn, LookIn:=xlValues, lookat:=xlWhole)
    If searchrng Is Nothing Then MsgBox wcn & "not found in ECL Student Information.": Exit Sub
    Worksheets("2496").Range("d3").Value = searchrng.Offset(0, 1).Value
    Worksheets("2496").Range("b5").Value = searchrng.Offset(0, 11).Value
    'Worksheets("2496").Range("b9").Value = searchrng.Offset(0, 15).Value
    Worksheets("2496").Range("d9").Value = searchrng.Offset(0, 18).Value
End Sub
Sub graduated(wcn As Integer)
    lastrow = Worksheets("Graduated").Cells(Rows.Count, "b").End(xlUp).Row
    Set searchGrad = Worksheets("Graduated").Range("b2:b" & lastrow).Find(wcn, LookIn:=xlValues, lookat:=xlWhole)
    If searchGrad Is Nothing Then MsgBox wcn & "not found in Graduated.": Exit Sub
    Worksheets("2496").Range("d5").Value = searchGrad.Offset(0, 6).Value
    Worksheets("2496").Range("d7").Value = searchGrad.Offset(0, 7).Value
    Worksheets("2496").Range("b9").Value = searchGrad.Offset(0, 8).Value
End Sub
Sub progress(wcn As Integer)
    lastrow = Worksheets("Progress").Cells(Rows.Count, "a").End(xlUp).Row
    Set searchProg = Worksheets("Progress").Range("a2:a" & lastrow).Find(wcn, LookIn:=xlValues, lookat:=xlWhole)
    If searchProg Is Nothing Then MsgBox wcn & "not found in Progress.": Exit Sub
    Worksheets("2496").Range("b7").Value = searchProg.Offset(0, 3).Value
End Sub
Sub alcpt(wcn As Integer)
    lastrow = Worksheets("ALCPT Scores").Cells(Rows.Count, "a").End(xlUp).Row
    Set searchALCPT = Worksheets("ALCPT Scores").Range("a2:a" & lastrow).Find(wcn, LookIn:=xlValues, lookat:=xlWhole)
    If searchALCPT Is Nothing Then MsgBox wcn & "not found in ALCPT Scores.": Exit Sub
    Worksheets("2496").Range("b11").Value = searchALCPT.Offset(0, 3).Value
End Sub
Sub ecl(wcn As Integer)
    lastrow = Worksheets("ECL Scores").Cells(Rows.Count, "a").End(xlUp).Row
    Set searchECL = Worksheets("ECL Scores").Range("a2:a" & lastrow).Find(wcn, LookIn:=xlValues, lookat:=xlWhole)
    If searchECL Is Nothing Then MsgBox wcn & "not found in ECL Scores.": Exit Sub
    Worksheets("2496").Range("d11").Value = searchECL.Offset(0, 2).Value
End Sub
Sub disc(wcn As Integer)
    Dim disc As String
    Dim pts As Integer
    Dim totalPts As Integer
    For Each iCell In Worksheets(" Discrepancy Log").Range("a:a").Cells
        If iCell.Value = wcn Then
            disc = disc & ", " & iCell.Offset(0, 3).Value
            pts = iCell.Offset(0, 5).Value
            If pts = 0 Then
                pts = 1
                totalPts = totalPts + pts
            Else
                totalPts = totalPts + pts
            End If
        End If
    Next iCell
    Worksheets("2496").Range("a15").Value = disc
    Worksheets("2496").Range("d14").Value = totalPts
End Sub
Sub excellent(wcn As Integer)
    Dim exc As String
    Dim pts As Integer
    Dim totalPts As Integer
    For Each iCell In Worksheets("Excellence").Range("a:a").Cells
        If iCell.Value = wcn Then
            exc = exc & ", " & iCell.Offset(0, 3).Value
            pts = iCell.Offset(0, 5).Value
            If pts = 0 Then
                pts = 1
                totalPts = totalPts + pts
            Else
                totalPts = totalPts + pts
            End If
        End If
    Next iCell
    Worksheets("2496").Range("a30").Value = exc
    Worksheets("2496").Range("d29").Value = totalPts
End Sub
Sub GetInfo()
    Dim wcn As Integer
    Dim searchrng As Range
    Dim searchGrad As Range
    Dim searchProg As Range
    Dim searchALCPT As Range
    Dim searchECL As Range
    Dim searchD As Range
    Dim searchE As Range
    Dim loc As Integer
    Dim lastrow As Long
    Dim dod As Long
    wcn = Worksheets("2496").Range("b3").Value
    Call studentInfo(wcn)
    Call graduated(wcn)
    Call progress(wcn)
    Call alcpt(wcn)
    Call ecl(wcn)
    Call disc(wcn)
    Call excellent(wcn)
End Sub
