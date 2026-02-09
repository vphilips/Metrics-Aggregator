Attribute VB_Name = "modFormHelpers"
Option Explicit

' ==================================================================================
' SHARED FORM HELPERS
' ==================================================================================
' These functions are used by the dynamically generated forms.
' ==================================================================================

Public Function CheckSecurityAccess() As Boolean
    Dim vbp As Object
    On Error Resume Next
    Set vbp = ThisWorkbook.VBProject
    On Error GoTo 0
    If vbp Is Nothing Or vbp.Protection = 1 Then
        MsgBox "Please enable 'Trust access to the VBA project object model' and unlock project.", vbCritical
        CheckSecurityAccess = False
    Else
        CheckSecurityAccess = True
    End If
End Function

Public Function NullIfEmpty(val As Variant) As Variant
    If IsNull(val) Or Trim(val & "") = "" Then NullIfEmpty = Null Else NullIfEmpty = val
End Function

Public Function SafeStr(val As Variant) As String
    If IsError(val) Then SafeStr = "" Else SafeStr = Trim(val & "")
End Function

Public Function GetSelectedItems(lst As Object) As String
    Dim i As Long, s As String: s = ""
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then s = s & lst.List(i) & ", "
    Next i
    If Len(s) > 0 Then s = Left(s, Len(s) - 2)
    GetSelectedItems = s
End Function

' --- LOOKUPS ---

Public Function GetMetricFromVariable(varName As String) As String
    On Error Resume Next
    Dim ws As Worksheet, rng As Range, f As Range
    Set ws = ThisWorkbook.Sheets("variable_metric_map")
    If ws Is Nothing Then GetMetricFromVariable = varName: Exit Function
    
    ' Assumes Variable in Col A, Metric in Col B
    Set rng = ws.Range("A:A")
    Set f = rng.Find(What:=Trim(varName), LookIn:=xlValues, LookAt:=xlWhole)
    If Not f Is Nothing Then
        GetMetricFromVariable = SafeStr(f.Offset(0, 1).Value)
    Else
        MsgBox "Variable '" & varName & "' not found in 'variable_metric_map' sheet.", vbCritical
        GetMetricFromVariable = ""
    End If
End Function

Public Function GetAccountKeys(lst As Object) As String
    ' 1. Load Account Map (Name -> GlobalID) into Dictionary for speed
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' TextCompare
    
    Dim ws As Worksheet, arr As Variant, i As Long, lastRow As Long
    Set ws = ThisWorkbook.Sheets("database")
    If ws Is Nothing Then MsgBox "Database sheet missing!": Exit Function
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow >= 2 Then
        ' Read Columns A through C
        arr = ws.Range("A2:C" & lastRow).Value ' A=Name, B=Key, C=GlobalID
        For i = 1 To UBound(arr, 1)
            If SafeStr(arr(i, 1)) <> "" Then
                ' Map Name (Col 1) to GlobalID (Col 3)
                Dim nKey As String: nKey = SafeStr(arr(i, 1))
                If Not dict.Exists(nKey) Then dict.Add nKey, arr(i, 3)
            End If
        Next i
    End If
    
    ' 2. Map Selected items
    Dim s As String, key As Variant, accName As String
    s = ""
    Dim missingList As String: missingList = ""
    
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            accName = SafeStr(lst.List(i))
            If dict.Exists(accName) Then
                s = s & dict(accName) & ","
            Else
                missingList = missingList & vbCrLf & "- " & accName
            End If
        End If
    Next i
    
    If Len(missingList) > 0 Then
        MsgBox "The following accounts could not be mapped to an Investor ID (check 'database' sheet):" & vbCrLf & missingList, vbCritical
        GetAccountKeys = ""
        Exit Function
    End If
    
    If Len(s) > 0 Then s = Left(s, Len(s) - 1)
    GetAccountKeys = s
End Function

' --- DATES ---

Public Function ParseDateForDB(val As Variant) As Variant
    If IsNull(val) Or Trim(val & "") = "" Then
        ParseDateForDB = Null
    Else
        Dim s As String
        s = Trim(val)
        s = Replace(s, "'", "")
        
        If IsDate(s) Then
            ParseDateForDB = CDate(s)
        ElseIf IsNumeric(s) Then
            On Error Resume Next
            ParseDateForDB = CDate(CDbl(s))
            If Err.Number <> 0 Then ParseDateForDB = Null
            On Error GoTo 0
        Else
            ParseDateForDB = Null
        End If
    End If
End Function

Public Function ParseDateToInt(val As Variant) As Variant
    If IsNull(val) Or Trim(val & "") = "" Then
        ParseDateToInt = Null
    Else
        Dim s As String
        s = Trim(val)
        If IsDate(s) Then
            ParseDateToInt = CLng(Format(CDate(s), "yyyymmdd"))
        ElseIf IsNumeric(s) And Len(s) = 8 Then
            ParseDateToInt = CLng(s)
        Else
            ParseDateToInt = Null
        End If
    End If
End Function

' --- CURRENCY ---

Public Function GetCurrencyID(ccyName As String) As Long
    On Error Resume Next
    Dim ws As Worksheet, rng As Range, f As Range
    Set ws = ThisWorkbook.Sheets("ccy_map")
    If ws Is Nothing Then GetCurrencyID = 0: Exit Function
    
    ' Look for Currency Code in Col A
    Set rng = ws.Range("A:A")
    Set f = rng.Find(What:=Trim(ccyName), LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not f Is Nothing Then
        GetCurrencyID = CLng(f.Offset(0, 1).Value) ' Col B has ID
    Else
        GetCurrencyID = 0 ' Not found
    End If
End Function
