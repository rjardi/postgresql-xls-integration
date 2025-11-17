' Archivo: get_stock_data_sem.vba
' Funciones UDF semanales que devuelven valores únicos desde PostgreSQL

Option Explicit

' Variables globales para conexión persistente
Private GlobalConnSem As Object
Private IsGlobalConnOpenSem As Boolean

' Obtener conexión global semanal
Private Function GetGlobalConnectionSem() As Object
    On Error GoTo ErrorHandler
    
    If GlobalConnSem Is Nothing Or Not IsGlobalConnOpenSem Then
        Set GlobalConnSem = CreateObject("ADODB.Connection")
        GlobalConnSem.Open "DSN=PostgreSQL35W"
        IsGlobalConnOpenSem = True
        Debug.Print "GetGlobalConnectionSem - Nueva conexión global creada"
    End If
    
    Set GetGlobalConnectionSem = GlobalConnSem
    Exit Function
    
ErrorHandler:
    Set GetGlobalConnectionSem = Nothing
    Debug.Print "GetGlobalConnectionSem - Error: " & Err.Description
End Function

' Cerrar conexión global (opcional)
Public Sub CloseGlobalConnectionSem()
    On Error Resume Next
    If Not GlobalConnSem Is Nothing And IsGlobalConnOpenSem Then
        GlobalConnSem.Close
        Set GlobalConnSem = Nothing
        IsGlobalConnOpenSem = False
        Debug.Print "CloseGlobalConnectionSem - Conexión cerrada"
    End If
    On Error GoTo 0
End Sub

' Ejecutar la stored procedure semanal api_xls.f_pla_get_data_stock_sem_v1
Private Function ExecuteGetDataStockSem( _
    ByVal p_unidad_operacional As Variant, _
    ByVal p_peticion As Variant, _
    ByVal p_producto_venta As Variant, _
    ByVal p_year As Variant, _
    ByVal p_week As Variant, _
    ByVal p_DiaVida_inicial As Variant, _
    ByVal p_DiaVida_final As Variant, _
    ByVal p_peso_inicial As Variant, _
    ByVal p_peso_final As Variant) As Variant
    
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    Dim sqlQuery As String
    
    Set cmd = Nothing
    Set rs = Nothing
    
    sqlQuery = "SELECT api_xls.f_pla_get_data_stock_sem_v1(?, ?, ?, ?, ?, ?, ?, ?, ?)"
    
    On Error GoTo ErrorHandler
    
    Set conn = GetGlobalConnectionSem()
    If conn Is Nothing Then
        ExecuteGetDataStockSem = "Error: No se pudo establecer conexión"
        Debug.Print "ExecuteGetDataStockSem - Error: No se pudo establecer conexión"
        Exit Function
    End If
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sqlQuery
    cmd.CommandType = 1 ' adCmdText

    cmd.Parameters.Append cmd.CreateParameter("p_unidad_operacional", 200, 1, 255, p_unidad_operacional)
    cmd.Parameters.Append cmd.CreateParameter("p_peticion", 200, 1, 255, p_peticion)
    cmd.Parameters.Append cmd.CreateParameter("p_producto_venta", 200, 1, 255, p_producto_venta)
    cmd.Parameters.Append cmd.CreateParameter("p_year", 3, 1, , CLng(p_year)) ' adInteger
    cmd.Parameters.Append cmd.CreateParameter("p_week", 3, 1, , CLng(p_week)) ' adInteger
    cmd.Parameters.Append cmd.CreateParameter("p_DiaVida_inicial", 3, 1, , CLng(p_DiaVida_inicial))
    cmd.Parameters.Append cmd.CreateParameter("p_DiaVida_final", 3, 1, , CLng(p_DiaVida_final))
    cmd.Parameters.Append cmd.CreateParameter("p_peso_inicial", 5, 1, , CDbl(p_peso_inicial)) ' adDouble
    cmd.Parameters.Append cmd.CreateParameter("p_peso_final", 5, 1, , CDbl(p_peso_final))
    
    Set rs = cmd.Execute
    
    If Not rs.EOF Then
        ExecuteGetDataStockSem = rs.Fields(0).value
    Else
        ExecuteGetDataStockSem = 0
    End If
    
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    On Error GoTo 0
    
    Exit Function
    
ErrorHandler:
    ExecuteGetDataStockSem = "Error: " & Err.Description & " (Error #" & Err.Number & ")"
    Debug.Print "ExecuteGetDataStockSem - Error: " & Err.Description & " (Error #" & Err.Number & ")"
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    On Error GoTo 0
End Function

' =============================================================================
' FUNCIONES PÚBLICAS SEMANALES (VALOR ÚNICO)
' =============================================================================

' GetStockAviSem - Métricas de stock semanales (QTY/KG/PM)
Public Function GetStockAviSem(p_unidad_operacional As String, p_peticion As String, p_producto_venta As String, p_year As Integer, p_week As Integer, p_DiaVida_inicial As Integer, p_DiaVida_final As Integer, p_peso_inicial As Double, p_peso_final As Double) As Variant
    On Error GoTo ErrorHandler
    
    GetStockAviSem = ExecuteGetDataStockSem( _
        p_unidad_operacional, _
        p_peticion, _
        p_producto_venta, _
        p_year, _
        p_week, _
        p_DiaVida_inicial, _
        p_DiaVida_final, _
        p_peso_inicial, _
        p_peso_final)
    Exit Function
    
ErrorHandler:
    GetStockAviSem = "Error: " & Err.Description & " (Error #" & Err.Number & ")"
    Debug.Print "GetStockAviSem - Error: " & Err.Description & " (Error #" & Err.Number & ")"
End Function

' GetEntradaAviSem - Entradas acumuladas en una semana
Public Function GetEntradaAviSem(p_unidad_operacional As String, p_peticion As String, p_producto_venta As String, p_year As Integer, p_week As Integer) As Variant
    On Error GoTo ErrorHandler
    
    GetEntradaAviSem = ExecuteGetDataStockSem( _
        p_unidad_operacional, _
        p_peticion, _
        p_producto_venta, _
        p_year, _
        p_week, _
        0, _
        9999, _
        0, _
        99.999)
    Exit Function
    
ErrorHandler:
    GetEntradaAviSem = "Error: " & Err.Description & " (Error #" & Err.Number & ")"
    Debug.Print "GetEntradaAviSem - Error: " & Err.Description & " (Error #" & Err.Number & ")"
End Function

' GetSalidasAviSem - Salidas acumuladas en una semana
Public Function GetSalidasAviSem(p_unidad_operacional As String, p_peticion As String, p_producto_venta As String, p_year As Integer, p_week As Integer, p_DiaVida_inicial As Integer, p_DiaVida_final As Integer, p_peso_inicial As Double, p_peso_final As Double) As Variant
    On Error GoTo ErrorHandler
    
    GetSalidasAviSem = ExecuteGetDataStockSem( _
        p_unidad_operacional, _
        p_peticion, _
        p_producto_venta, _
        p_year, _
        p_week, _
        p_DiaVida_inicial, _
        p_DiaVida_final, _
        p_peso_inicial, _
        p_peso_final)
    Exit Function
    
ErrorHandler:
    GetSalidasAviSem = "Error: " & Err.Description & " (Error #" & Err.Number & ")"
    Debug.Print "GetSalidasAviSem - Error: " & Err.Description & " (Error #" & Err.Number & ")"
End Function



