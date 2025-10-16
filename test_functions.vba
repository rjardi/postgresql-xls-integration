Private Sub Test_GetEntradaAvi()
    ' Cerrar conexión global primero
    CloseGlobalConnection
    Debug.Print ("")
    Debug.Print ("Test_GetEntradaAvi-----------------------------------")
    test_value = GetEntradaAvi("TOR", "ENTRADAS_QTY", "BRAM", "2025-10-01", "2025-10-10")
    
    Debug.Print (test_value)
    CloseGlobalConnection ' Cerrar conexión
    Debug.Print ("")
End Sub
Private Sub Test_GetEntradaAvi_2()
    ' Cerrar conexión global primero
    CloseGlobalConnection
    Debug.Print ("")
    Debug.Print ("Test_GetEntradaAvi_2-----------------------------------")
    test_value = GetEntradaAvi("PDX", "ENTRADAS_QTY", "BRAM", "2025-10-01", "2025-10-10")
    Debug.Print (test_value)
    CloseGlobalConnection ' Cerrar conexión
    Debug.Print ("")
End Sub

Private Sub Test_GetEntradaAvi_Debug()
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    Dim sqlQuery As String
    
    ' Crear conexión nueva para debug
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "DSN=PostgreSQL35W"
    
    ' Query directa sin parámetros preparados
    sqlQuery = "SELECT api_xls.f_pla_get_data_stock('TOR', 'ENTRADAS_QTY', 'BRAM', '2025-10-16', '2025-10-01', '2025-10-10', 0, 9999, 0, 99.999)"
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sqlQuery
    cmd.CommandType = 1
    
    Set rs = cmd.Execute
    
    If Not rs.EOF Then
        Debug.Print "SQL DIRECTO TOR: " & rs.Fields(0).Value
    End If
    
    ' Limpiar
    rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    conn.Close
    Set conn = Nothing
End Sub

Private Sub Test_Characters()
    Dim tor As String
    tor = "TOR"
    Debug.Print "Longitud TOR: " & Len(tor)
    Debug.Print "Códigos ASCII: " & Asc(Left(tor, 1)) & ", " & Asc(Mid(tor, 2, 1)) & ", " & Asc(Right(tor, 1))
End Sub

Private Sub Test_OtherValues()
    ' Probar con otros valores similares
    Debug.Print "Test 1: " & GetEntradaAvi("TOR ", "ENTRADAS_QTY", "BRAM", "2025-10-01", "2025-10-10") ' con espacio
    Debug.Print "Test 2: " & GetEntradaAvi("tor", "ENTRADAS_QTY", "BRAM", "2025-10-01", "2025-10-10") ' minúsculas
    Debug.Print "Test 3: " & GetEntradaAvi("TOR", "entradas_qty", "BRAM", "2025-10-01", "2025-10-10") ' minúsculas petición
End Sub

Private Sub Test_TOR_Debug_Complete()
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    Dim sqlQuery As String
    
    ' Crear conexión nueva
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "DSN=PostgreSQL35W"
    
    ' Query directa con TOR
    sqlQuery = "SELECT api_xls.f_pla_get_data_stock('TOR', 'ENTRADAS_QTY', 'BRAM', '2025-10-16', '2025-10-01', '2025-10-10', 0, 9999, 0, 99.999)"
    
    Debug.Print "=== QUERY TOR ==="
    Debug.Print sqlQuery
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sqlQuery
    cmd.CommandType = 1
    
    On Error GoTo ErrorHandler
    Set rs = cmd.Execute
    
    If Not rs.EOF Then
        Debug.Print "TOR resultado: " & rs.Fields(0).Value
        Debug.Print "Tipo de resultado: " & TypeName(rs.Fields(0).Value)
    Else
        Debug.Print "TOR: No hay datos"
    End If
    
    ' Limpiar
    rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    conn.Close
    Set conn = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR TOR: " & Err.Description & " (Error #" & Err.Number & ")"
    ' Limpiar en caso de error
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    conn.Close
    Set conn = Nothing
End Sub


Sub Test_PDX_Debug_Complete()
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    Dim sqlQuery As String
    
    ' Crear conexión nueva
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "DSN=PostgreSQL35W"
    
    ' Query directa con PDX
    sqlQuery = "SELECT api_xls.f_pla_get_data_stock('PDX', 'ENTRADAS_QTY', 'BRAM', '2025-10-16', '2025-10-01', '2025-10-10', 0, 9999, 0, 99.999)"
    
    Debug.Print "=== QUERY PDX ==="
    Debug.Print sqlQuery
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sqlQuery
    cmd.CommandType = 1
    
    On Error GoTo ErrorHandler
    Set rs = cmd.Execute
    
    If Not rs.EOF Then
        Debug.Print "PDX resultado: " & rs.Fields(0).Value
        Debug.Print "Tipo de resultado: " & TypeName(rs.Fields(0).Value)
    Else
        Debug.Print "PDX: No hay datos"
    End If
    
    ' Limpiar
    rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    conn.Close
    Set conn = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR PDX: " & Err.Description & " (Error #" & Err.Number & ")"
    ' Limpiar en caso de error
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    conn.Close
    Set conn = Nothing
End Sub


Sub Test_Simple_Query()
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    Dim sqlQuery As String
    
    ' Crear conexión nueva
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "DSN=PostgreSQL35W"
    
    ' Query simple para verificar conexión
    sqlQuery = "SELECT 'TOR' as test_value"
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sqlQuery
    cmd.CommandType = 1
    
    Set rs = cmd.Execute
    
    If Not rs.EOF Then
        Debug.Print "Query simple: " & rs.Fields(0).Value
    End If
    
    ' Limpiar
    rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    conn.Close
    Set conn = Nothing
End Sub


