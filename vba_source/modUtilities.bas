Attribute VB_Name = "modUtilities"
Option Explicit

Public Sub OutputToResults(rs As Object)
    Dim ws As Worksheet
    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    ' 1. Handle "Results" Sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Results")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Create new if missing
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Results"
    Else
        ' Clear existing
        ws.Cells.Clear
    End If
    
    ws.Activate
    
    ' 2. Validation Checks
    If rs Is Nothing Then
         Application.ScreenUpdating = True
         Exit Sub ' Error already handled in ExecuteSP
    End If
    
    If rs.State = 0 Then
        MsgBox "Recordset is closed. No data returned.", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    If rs.EOF And rs.BOF Then
        MsgBox "Query executed successfully but returned zero records.", vbInformation
        rs.Close
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' 3. Write Headers
    For i = 0 To rs.Fields.Count - 1
        ws.Cells(1, i + 1).Value = rs.Fields(i).Name
    Next i
    
    ' 4. Style Headers
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, rs.Fields.Count))
        .Font.Bold = True
        .Interior.Color = RGB(230, 230, 230)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    ' 5. Dump Data
    ws.Cells(2, 1).CopyFromRecordset rs
    
    ' 6. AutoFit
    ws.Columns.AutoFit
    
    ' Cleanup
    rs.Close
    Set rs = Nothing
    
    Application.ScreenUpdating = True
    MsgBox "Data retrieved and populated in 'Results' sheet.", vbInformation
End Sub

Public Sub BuildParams(ByRef pNames As Variant, ByRef pValues As Variant, ParamArray args() As Variant)
    ' Helper to build arrays from Name/Value pairs
    ' Usage: BuildParams n, v, "@Name1", val1, "@Name2", val2
    Dim n As Integer
    Dim i As Integer, k As Integer
    
    If UBound(args) < 0 Then Exit Sub
    
    n = (UBound(args) + 1) \ 2
    ReDim pNames(0 To n - 1)
    ReDim pValues(0 To n - 1)
    
    k = 0
    For i = 0 To UBound(args) Step 2
        pNames(k) = args(i)
        pValues(k) = args(i + 1)
        k = k + 1
    Next i
End Sub

' Centralized Validation logic
Public Function ValidateMandatory(ctrl As Object, fieldName As String) As Boolean
    If Len(Trim(ctrl.Value)) = 0 Then
        MsgBox fieldName & " is a mandatory field.", vbExclamation, "Missing Input"
        ctrl.SetFocus
        ValidateMandatory = False
    Else
        ValidateMandatory = True
    End If
End Function

Public Function ValidateDate(ctrl As Object, fieldName As String) As Boolean
    If Len(Trim(ctrl.Value)) > 0 Then
        If Not IsDate(ctrl.Value) Then
            MsgBox fieldName & " must be a valid date.", vbExclamation, "Invalid Input"
            ctrl.SetFocus
            ValidateDate = False
            Exit Function
        End If
    End If
    ValidateDate = True
End Function
