' Archivo: get_stock_data_grj.vba
' Función UDF para obtener datos por granja/lote/nave/art desde PostgreSQL

Option Explicit

' Variables globales para conexión persistente
Private GlobalConnGrj As Object
Private IsGlobalConnOpenGrj As Boolean

' Obtener conexión global
Private Function GetGlobalConnectionGrj() As Object
    On Error GoTo ErrorHandler
    
    If GlobalConnGrj Is Nothing Or Not IsGlobalConnOpenGrj Then
        Set GlobalConnGrj = CreateObject("ADODB.Connection")
        GlobalConnGrj.Open "DSN=PostgreSQL35W"
        IsGlobalConnOpenGrj = True
        Debug.Print "GetGlobalConnectionGrj - Nueva conexión global creada"
    End If
    
    Set GetGlobalConnectionGrj = GlobalConnGrj
    Exit Function
    
ErrorHandler:
    Set GetGlobalConnectionGrj = Nothing
    Debug.Print "GetGlobalConnectionGrj - Error: " & Err.Description
End Function

' Cerrar conexión global (opcional)
Public Sub CloseGlobalConnectionGrj()
    On Error Resume Next
    If Not GlobalConnGrj Is Nothing And IsGlobalConnOpenGrj Then
        GlobalConnGrj.Close
        Set GlobalConnGrj = Nothing
        IsGlobalConnOpenGrj = False
        Debug.Print "CloseGlobalConnectionGrj - Conexión cerrada"
    End If
    On Error GoTo 0
End Sub

' Ejecutar la stored procedure api_xls.f_pla_get_data_stock_grj_v1
Private Function ExecuteGetDataStockGrj( _
    ByVal p_unidad_operacional As Variant, _
    ByVal p_peticion As Variant, _
    ByVal p_fecha_dato As Variant, _
    ByVal p_granja As Variant, _
    ByVal p_lote As Variant, _
    ByVal p_nave As Variant, _
    ByVal p_articulo As Variant) As Variant
    
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    Dim sqlQuery As String
    
    Set cmd = Nothing
    Set rs = Nothing
    
    sqlQuery = "SELECT api_xls.f_pla_get_data_stock_grj_v1(?, ?, ?, ?, ?, ?, ?)"
    
    On Error GoTo ErrorHandler
    
    Set conn = GetGlobalConnectionGrj()
    If conn Is Nothing Then
        ExecuteGetDataStockGrj = "Error: No se pudo establecer conexión"
        Debug.Print "ExecuteGetDataStockGrj - Error: No se pudo establecer conexión"
        Exit Function
    End If
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sqlQuery
    cmd.CommandType = 1 ' adCmdText
    
    cmd.Parameters.Append cmd.CreateParameter("p_unidad_operacional", 200, 1, 255, p_unidad_operacional)
    cmd.Parameters.Append cmd.CreateParameter("p_peticion", 200, 1, 255, p_peticion)
    cmd.Parameters.Append cmd.CreateParameter("p_fecha_dato", 200, 1, 255, Format(p_fecha_dato, "yyyy-mm-dd")) ' adVarChar como string
    cmd.Parameters.Append cmd.CreateParameter("p_granja", 200, 1, 255, p_granja)
    cmd.Parameters.Append cmd.CreateParameter("p_lote", 200, 1, 255, p_lote)
    cmd.Parameters.Append cmd.CreateParameter("p_nave", 200, 1, 255, p_nave)
    cmd.Parameters.Append cmd.CreateParameter("p_articulo", 200, 1, 255, p_articulo)
    
    Set rs = cmd.Execute
    
    If Not rs.EOF Then
        ExecuteGetDataStockGrj = rs.Fields(0).value
    Else
        ExecuteGetDataStockGrj = 0
    End If
    
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    On Error GoTo 0
    
    Exit Function
    
ErrorHandler:
    ExecuteGetDataStockGrj = "Error: " & Err.Description & " (Error #" & Err.Number & ")"
    Debug.Print "ExecuteGetDataStockGrj - Error: " & Err.Description & " (Error #" & Err.Number & ")"
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    On Error GoTo 0
End Function

' =============================================================================
' FUNCIÓN PÚBLICA PARA EXCEL
' =============================================================================

Public Function GetStockAviGrj(p_unidad_operacional As String, p_peticion As String, p_fecha_dato As Date, p_granja As String, p_lote As String, p_nave As String, p_articulo As String) As Variant
    On Error GoTo ErrorHandler
    
    GetStockAviGrj = ExecuteGetDataStockGrj( _
        p_unidad_operacional, _
        p_peticion, _
        p_fecha_dato, _
        p_granja, _
        p_lote, _
        p_nave, _
        p_articulo)
    Exit Function
    
ErrorHandler:
    GetStockAviGrj = "Error: " & Err.Description & " (Error #" & Err.Number & ")"
    Debug.Print "GetStockAviGrj - Error: " & Err.Description & " (Error #" & Err.Number & ")"
End Function



