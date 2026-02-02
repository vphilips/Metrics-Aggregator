Attribute VB_Name = "modDatabase"
Option Explicit

' ==========================================
' CONFIGURATION
' ==========================================
' UPDATE THIS STRING WITH YOUR SPECIFIC SERVER DETAILS
' NOTE: On Mac, standard SQLOLEDB Providers do not work. You may need specific ODBC drivers.
Private Const DB_CONNECTION_STRING As String = "Provider=SQLOLEDB;Data Source=My_Server;Initial Catalog=My_Database;Integrated Security=SSPI;"

' ==========================================
' PUBLIC FUNCTIONS
' ==========================================

Public Function GetConnection() As Object
    Dim conn As Object
    On Error Resume Next
    Set conn = CreateObject("ADODB.Connection")
    On Error GoTo ConnError
    
    If conn Is Nothing Then
        MsgBox "Could not create ADODB.Connection object. If you are on Mac, ensure you have necessary drivers or use a compatible method.", vbCritical
        Exit Function
    End If
    
    conn.ConnectionString = DB_CONNECTION_STRING
    conn.Open
    Set GetConnection = conn
    Exit Function

ConnError:
    MsgBox "Error connecting to database: " & Err.Description, vbCritical, "Database Connection Error"
    Set GetConnection = Nothing
End Function

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
