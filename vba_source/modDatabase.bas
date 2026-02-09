Attribute VB_Name = "modDatabase"
Option Explicit

' ==========================================
' CONFIGURATION
' ==========================================
' Connection String Configuration
' You can use a trusted connection or a specific user.
' Examples:
' SQL Server (Standard): "Provider=SQLOLEDB;Data Source=SERVER_NAME;Initial Catalog=DB_NAME;Integrated Security=SSPI;"
' SQL Server (Native Client): "Provider=SQLNCLI11;Data Source=SERVER_NAME;Initial Catalog=DB_NAME;Integrated Security=SSPI;"
' ODBC (Mac/Windows): "Driver={SQL Server};Server=SERVER_NAME;Database=DB_NAME;Trusted_Connection=yes;"

Private Const DB_CONN_STR_MODERN As String = "Provider=MSOLEDBSQL;Data Source=My_Server;Initial Catalog=My_Database;Integrated Security=SSPI;"
Private Const DB_CONN_STR_LEGACY As String = "Provider=SQLOLEDB;Data Source=My_Server;Initial Catalog=My_Database;Integrated Security=SSPI;"

' ==========================================
' PUBLIC FUNCTIONS
' ==========================================

Public Function GetConnection() As Object
    Dim conn As Object
    On Error Resume Next
    Set conn = CreateObject("ADODB.Connection")
    
    If conn Is Nothing Then
        MsgBox "Could not create ADODB.Connection object.", vbCritical
        Exit Function
    End If
    
    ' Attempt 1: Modern Driver
    On Error Resume Next
    conn.ConnectionString = DB_CONN_STR_MODERN
    conn.Open
    
    If conn.State = 1 Then
        Set GetConnection = conn
        Exit Function
    End If
    
    ' Attempt 2: Legacy Driver (Fallback)
    Err.Clear
    conn.ConnectionString = DB_CONN_STR_LEGACY
    conn.Open
    
    If conn.State = 1 Then
        Set GetConnection = conn
    Else
        MsgBox "Error connecting to database." & vbCrLf & _
               "Tried MSOLEDBSQL and SQLOLEDB." & vbCrLf & _
               "Last Error: " & Err.Description, vbCritical, "Database Connection Error"
        Set GetConnection = Nothing
    End If
    On Error GoTo 0
End Function

Public Sub CloseConnection(conn As Object)
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    Set conn = Nothing
    On Error GoTo 0
End Sub

Public Function ExecuteSP(spName As String, pNames As Variant, pValues As Variant) As Object
    ' param: pNames Array of Strings
    ' param: pValues Array of Values
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    Dim i As Integer
    
    Set conn = GetConnection()
    If conn Is Nothing Then Exit Function
    
    Set cmd = CreateObject("ADODB.Command")
    
    On Error GoTo CmdError
    With cmd
        .ActiveConnection = conn
        .CommandText = spName
        .CommandType = 4 ' adCmdStoredProc
        
        .Parameters.Refresh
        
        If IsArray(pNames) Then
            For i = LBound(pNames) To UBound(pNames)
                Dim paramName As String
                paramName = pNames(i)
                
                ' Find param in collection
                Dim p As Object
                Dim paramExists As Boolean
                paramExists = False
                
                For Each p In .Parameters
                    If LCase(p.Name) = LCase(paramName) Then
                        p.Value = pValues(i)
                        paramExists = True
                        Exit For
                    End If
                Next p
                
                If Not paramExists Then
                    Debug.Print "Warning: Parameter " & paramName & " passed but not found in SP " & spName
                End If
            Next i
        End If
    End With
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' adUseClient
    rs.Open cmd
    
    Set cmd.ActiveConnection = Nothing
    conn.Close
    
    Set ExecuteSP = rs
    Exit Function

CmdError:
    MsgBox "Error executing Stored Procedure: " & Err.Description, vbCritical, "Database Execution Error"
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    Set ExecuteSP = Nothing
End Function
