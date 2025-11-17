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
        GlobalConn.Open "DSN=PostgreSQL35W"
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
    
    On Error GoTo ErrorHandler
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "DSN=PostgreSQL35W"
    
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
    
    On Error GoTo ErrorHandler
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "DSN=PostgreSQL35W"
    
    TestConnectionSSL = "Conexión SSL exitosa"
    Debug.Print "TestConnectionSSL - Conexión SSL exitosa"
    
    Exit Function
    
ErrorHandler:
    TestConnectionSSL = "Error SSL: " & Err.Description
    Debug.Print "TestConnectionSSL - Error SSL: " & Err.Description
End Function

Private Function ToSql(v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        ToSql = "NULL"
    ElseIf IsDate(v) Then
        ToSql = "'" & Format$(CDate(v), "yyyy-mm-dd") & "'"
    ElseIf IsNumeric(v) Then
        ' Convertir comas a puntos para números
        ToSql = Replace(CStr(v), ",", ".")
    Else
        ToSql = "'" & Replace(CStr(v), "'", "''") & "'"
    End If
End Function

' =============================================================================
' NUEVAS FUNCIONES PARA OBTENER DATOS DE STOCK
' =============================================================================

 ' Función auxiliar reutilizable para llamar a api_xls.f_pla_get_data_stock
 ' Firma fija con los 10 parámetros en orden exacto requerido por PostgreSQL
Private Function ExecuteGetDataStock( _
    ByVal p_unidad_operacional As Variant, _
    ByVal p_peticion As Variant, _
    ByVal p_producto_venta As Variant, _
    ByVal p_fecha_dato As Variant, _
    ByVal p_fch_inicial As Variant, _
    ByVal p_fch_final As Variant, _
    ByVal p_DiaVida_inicial As Variant, _
    ByVal p_DiaVida_final As Variant, _
    ByVal p_peso_inicial As Variant, _
    ByVal p_peso_final As Variant) As Variant
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    Dim sqlQuery As String
    
    ' Inicializar variables
    Set cmd = Nothing
    Set rs = Nothing
    
    ' Query SQL con los 10 parámetros fijos
    sqlQuery = "SELECT api_xls.f_pla_get_data_stock_v1(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    
    On Error GoTo ErrorHandler
    
    ' Obtener conexión global (reutiliza la existente)
    Set conn = GetGlobalConnection()
    If conn Is Nothing Then
        ExecuteGetDataStock = "Error: No se pudo establecer conexión"
        Debug.Print "ExecuteGetDataStock - Error: No se pudo establecer conexión"
        Exit Function
    End If
    
    ' Crear comando
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sqlQuery
    cmd.CommandType = 1 ' adCmdText
    
    ' Agregar los 10 parámetros en el orden exacto
    cmd.Parameters.Append cmd.CreateParameter("p_unidad_operacional", 200, 1, 255, p_unidad_operacional) ' adVarChar
    cmd.Parameters.Append cmd.CreateParameter("p_peticion", 200, 1, 255, p_peticion) ' adVarChar
    cmd.Parameters.Append cmd.CreateParameter("p_producto_venta", 200, 1, 255, p_producto_venta) ' adVarChar
    cmd.Parameters.Append cmd.CreateParameter("p_fecha_dato", 200, 1, 255, Format(p_fecha_dato, "yyyy-mm-dd")) ' adVarChar como string
    cmd.Parameters.Append cmd.CreateParameter("p_fch_inicial", 200, 1, 255, Format(p_fch_inicial, "yyyy-mm-dd")) ' adVarChar como string
    cmd.Parameters.Append cmd.CreateParameter("p_fch_final", 200, 1, 255, Format(p_fch_final, "yyyy-mm-dd")) ' adVarChar como string
    cmd.Parameters.Append cmd.CreateParameter("p_DiaVida_inicial", 200, 1, 255, CStr(p_DiaVida_inicial)) ' adVarChar como string
    cmd.Parameters.Append cmd.CreateParameter("p_DiaVida_final", 200, 1, 255, CStr(p_DiaVida_final)) ' adVarChar como string
    cmd.Parameters.Append cmd.CreateParameter("p_peso_inicial", 200, 1, 255, Replace(CStr(p_peso_inicial), ",", ".")) ' adVarChar con punto decimal
    cmd.Parameters.Append cmd.CreateParameter("p_peso_final", 200, 1, 255, Replace(CStr(p_peso_final), ",", ".")) ' adVarChar con punto decimal
        
    ' Ejecutar query
    Set rs = cmd.Execute
    
    Dim debugSql As String
    debugSql = "SELECT api_xls.f_pla_get_data_stock_v1(" & _
            ToSql(p_unidad_operacional) & ", " & _
            ToSql(p_peticion) & ", " & _
            ToSql(p_producto_venta) & ", " & _
            ToSql(p_fecha_dato) & ", " & _
            ToSql(p_fch_inicial) & ", " & _
            ToSql(p_fch_final) & ", " & _
            ToSql(p_DiaVida_inicial) & ", " & _
            ToSql(p_DiaVida_final) & ", " & _
            ToSql(p_peso_inicial) & ", " & _
            ToSql(p_peso_final) & ")"
    Debug.Print "SQL DEBUG => "; debugSql
    
    ' Obtener resultado
    If Not rs.EOF Then
        ExecuteGetDataStock = rs.Fields(0).Value
        
        Debug.Print "ExecuteGetDataStock - Resultado: " & ExecuteGetDataStock
    Else
        ExecuteGetDataStock = "No data found"
        Debug.Print "ExecuteGetDataStock - No data found"
    End If
    
    ' Limpiar solo el recordset y comando (NO la conexión)
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    On Error GoTo 0
    
    Exit Function
    
ErrorHandler:
    ExecuteGetDataStock = "Error: " & Err.Description & " (Error #" & Err.Number & ")"
    Debug.Print "ExecuteGetDataStock - Error: " & Err.Description & " (Error #" & Err.Number & ")"
    ' Limpiar en caso de error (NO la conexión global)
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    On Error GoTo 0
End Function

' Función GetStockAvi - Obtiene datos de stock con parámetros específicos
Public Function GetStockAvi(p_unidad_operacional As String, p_peticion As String, p_producto_venta As String, p_fecha_dato As Date, p_DiaVida_inicial As Integer, p_DiaVida_final As Integer, p_peso_inicial As Double, p_peso_final As Double) As Variant
    On Error GoTo ErrorHandler
    Dim fechaHoy As Date: fechaHoy = Date
    
    ' Mapear a los 10 parámetros: usar fecha_dato y defaults para rango fechas
    GetStockAvi = ExecuteGetDataStock( _
        p_unidad_operacional, _
        p_peticion, _
        p_producto_venta, _
        p_fecha_dato, _
        fechaHoy, _
        fechaHoy, _
        p_DiaVida_inicial, _
        p_DiaVida_final, _
        p_peso_inicial, _
        p_peso_final)
    
    Exit Function
    
ErrorHandler:
    GetStockAvi = "Error: " & Err.Description & " (Error #" & Err.Number & ")"
    Debug.Print "GetStockAvi - Error: " & Err.Description & " (Error #" & Err.Number & ")"
End Function

' Función GetEntradaAvi - Obtiene datos de entradas con parámetros específicos
Public Function GetEntradaAvi(p_unidad_operacional As String, p_peticion As String, p_producto_venta As String, p_fch_inicial As Date, p_fch_final As Date) As Variant
    On Error GoTo ErrorHandler
    Dim fechaHoy As Date: fechaHoy = Date
    
    ' Mapear a los 10 parámetros: usar rango fechas y defaults para resto
    GetEntradaAvi = ExecuteGetDataStock( _
        p_unidad_operacional, _
        p_peticion, _
        p_producto_venta, _
        fechaHoy, _
        p_fch_inicial, _
        p_fch_final, _
        0, _
        9999, _
        0, _
        99.999)
    
    Exit Function
    
ErrorHandler:
    GetEntradaAvi = "Error: " & Err.Description & " (Error #" & Err.Number & ")"
    Debug.Print "GetEntradaAvi - Error: " & Err.Description & " (Error #" & Err.Number & ")"
End Function

' Función GetSalidasAvi - Obtiene datos de salidas con parámetros específicos
Public Function GetSalidasAvi(p_unidad_operacional As String, p_peticion As String, p_producto_venta As String, p_fch_inicial As Date, p_fch_final As Date, p_DiaVida_inicial As Integer, p_DiaVida_final As Integer, p_peso_inicial As Double, p_peso_final As Double) As Variant
    On Error GoTo ErrorHandler
    Dim fechaHoy As Date: fechaHoy = Date
    
    ' Mapear a los 10 parámetros: usar rango fechas y filtros adicionales
    GetSalidasAvi = ExecuteGetDataStock( _
        p_unidad_operacional, _
        p_peticion, _
        p_producto_venta, _
        fechaHoy, _
        p_fch_inicial, _
        p_fch_final, _
        p_DiaVida_inicial, _
        p_DiaVida_final, _
        p_peso_inicial, _
        p_peso_final)
    
    Exit Function
    
ErrorHandler:
    GetSalidasAvi = "Error: " & Err.Description & " (Error #" & Err.Number & ")"
    Debug.Print "GetSalidasAvi - Error: " & Err.Description & " (Error #" & Err.Number & ")"
End Function



