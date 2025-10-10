' Archivo: get_stock_data.vba
' Función UDF para obtener stock desde PostgreSQL

' Variables globales para conexión persistente
Private GlobalConn As Object
Private IsGlobalConnOpen As Boolean

' Función para obtener conexión global
Private Function GetGlobalConnection() As Object
    On Error GoTo ErrorHandler
    
    ' Si no hay conexión o está cerrada, crear nueva
    If GlobalConn Is Nothing Or Not IsGlobalConnOpen Then
        Set GlobalConn = CreateObject("ADODB.Connection")
        GlobalConn.Open "FileDSN=" & ThisWorkbook.Path & "\postgresql.dsn"
        IsGlobalConnOpen = True
        Debug.Print "GetGlobalConnection - Nueva conexión global creada"
    End If
    
    Set GetGlobalConnection = GlobalConn
    Exit Function
    
ErrorHandler:
    Set GetGlobalConnection = Nothing
    Debug.Print "GetGlobalConnection - Error: " & Err.Description
End Function

' Función para cerrar conexión global
Public Sub CloseGlobalConnection()
    On Error Resume Next
    If Not GlobalConn Is Nothing And IsGlobalConnOpen Then
        GlobalConn.Close
        Set GlobalConn = Nothing
        IsGlobalConnOpen = False
        Debug.Print "CloseGlobalConnection - Conexión global cerrada"
    End If
    On Error GoTo 0
End Sub

Public Function GetStockData(p_empavi As String, p_erpcodave As String, p_fch As String) As Variant
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    Dim sqlQuery As String
    
    ' Inicializar variables
    Set cmd = Nothing
    Set rs = Nothing
    
    ' Query SQL
    sqlQuery = "SELECT api_xls.f_pla_qty_stock(?, ?, ?)"
    
    On Error GoTo ErrorHandler
    
    ' Obtener conexión global (reutiliza la existente)
    Set conn = GetGlobalConnection()
    If conn Is Nothing Then
        GetStockData = "Error: No se pudo establecer conexión"
        Debug.Print "GetStockData - Error: No se pudo establecer conexión"
        Exit Function
    End If
    
    ' Crear comando
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sqlQuery
    cmd.CommandType = 1 ' adCmdText
    
    ' Agregar parámetros
    cmd.Parameters.Append cmd.CreateParameter("p_empavi", 200, 1, 255, p_empavi)
    cmd.Parameters.Append cmd.CreateParameter("p_erpcodave", 200, 1, 255, p_erpcodave)
    cmd.Parameters.Append cmd.CreateParameter("p_fch", 200, 1, 255, p_fch)
    
    ' Ejecutar query
    Set rs = cmd.Execute
    
    ' Obtener resultado
    If Not rs.EOF Then
        GetStockData = rs.Fields(0).Value
        Debug.Print "GetStockData - Resultado: " & GetStockData
    Else
        GetStockData = "No data found"
        Debug.Print "GetStockData - No data found"
    End If
    
    ' Limpiar solo el recordset y comando (NO la conexión)
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    On Error GoTo 0
    
    Exit Function
    
ErrorHandler:
    GetStockData = "Error: " & Err.Description & " (Error #" & Err.Number & ")"
    Debug.Print "GetStockData - Error: " & Err.Description & " (Error #" & Err.Number & ")"
    ' Limpiar en caso de error (NO la conexión global)
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    On Error GoTo 0
End Function

' Función para inicializar conexión global
Public Sub InitializeConnection()
    Dim conn As Object
    Set conn = GetGlobalConnection()
    If Not conn Is Nothing Then
        Debug.Print "InitializeConnection - Conexión global inicializada"
    Else
        Debug.Print "InitializeConnection - Error al inicializar conexión"
    End If
End Sub

' Función para verificar estado de conexión global
Public Function GetConnectionStatus() As String
    If IsGlobalConnOpen And Not GlobalConn Is Nothing Then
        GetConnectionStatus = "Conexión global activa"
        Debug.Print "GetConnectionStatus - Conexión global activa"
    Else
        GetConnectionStatus = "Conexión global inactiva"
        Debug.Print "GetConnectionStatus - Conexión global inactiva"
    End If
End Function

' Función simple para probar conexión
Public Function TestConnection() As String
    Dim conn As Object
    Dim dsnPath As String
    
    dsnPath = ThisWorkbook.Path & "\postgresql.dsn"
    
    On Error GoTo ErrorHandler
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "FileDSN=" & dsnPath
    
    TestConnection = "Conexión exitosa"
    Debug.Print "TestConnection - Conexión exitosa"
    
    ' Cerrar esta conexión de prueba
    conn.Close
    Set conn = Nothing
    
    Exit Function
    
ErrorHandler:
    TestConnection = "Error: " & Err.Description
    Debug.Print "TestConnection - Error: " & Err.Description
End Function

' Función para probar con SSL usando DSN
Public Function TestConnectionSSL() As String
    Dim conn As Object
    Dim dsnPath As String
    
    dsnPath = ThisWorkbook.Path & "\postgresql.dsn"
    
    On Error GoTo ErrorHandler
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "FileDSN=" & dsnPath
    
    TestConnectionSSL = "Conexión SSL exitosa"
    Debug.Print "TestConnectionSSL - Conexión SSL exitosa"
    
    Exit Function
    
ErrorHandler:
    TestConnectionSSL = "Error SSL: " & Err.Description
    Debug.Print "TestConnectionSSL - Error SSL: " & Err.Description
End Function

