' Archivo: set_stock_data.vba
' Funciones UDF para obtener datos de stock desde PostgreSQL y pintar tablas en Excel

' Variables globales para conexión persistente (reutilizar las del archivo principal)
Private GlobalConn As Object
Private IsGlobalConnOpen As Boolean

' Función para obtener conexión global (reutilizar del archivo principal)
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

' Función auxiliar para parsear JSON manualmente (versión mejorada y robusta)
Private Function ParseJsonDataImproved(jsonString As String) As String
    On Error GoTo ErrorHandler
    
    ' Buscar el campo "data" de forma más robusta
    Dim dataStart As Long
    Dim dataEnd As Long
    Dim dataContent As String
    Dim bracketCount As Integer
    Dim i As Long
    
    ' Buscar el inicio del campo "data"
    dataStart = InStr(1, jsonString, """data""")
    If dataStart = 0 Then
        ParseJsonDataImproved = ""
        Exit Function
    End If
    
    ' Buscar el inicio del array [
    dataStart = InStr(dataStart, jsonString, "[")
    If dataStart = 0 Then
        ParseJsonDataImproved = ""
        Exit Function
    End If
    
    ' Encontrar el final del array contando corchetes
    bracketCount = 1
    i = dataStart + 1
    Do While i <= Len(jsonString) And bracketCount > 0
        If Mid(jsonString, i, 1) = "[" Then
            bracketCount = bracketCount + 1
        ElseIf Mid(jsonString, i, 1) = "]" Then
            bracketCount = bracketCount - 1
        End If
        i = i + 1
    Loop
    
    If bracketCount = 0 Then
        dataEnd = i - 1
        dataContent = Mid(jsonString, dataStart, dataEnd - dataStart + 1)
        ParseJsonDataImproved = dataContent
    Else
        ParseJsonDataImproved = ""
    End If
    
    Exit Function
    
ErrorHandler:
    ParseJsonDataImproved = ""
    Debug.Print "ParseJsonDataImproved - Error: " & Err.Description
End Function

' Función auxiliar para extraer claves de un objeto JSON (versión mejorada)
Private Function ExtractKeysImproved(jsonObject As String) As String()
    On Error GoTo ErrorHandler
    
    Dim keys() As String
    Dim keyCount As Integer
    Dim pos As Long
    Dim keyStart As Long
    Dim keyEnd As Long
    Dim key As String
    Dim char As String
    
    ' Inicializar array con tamaño suficiente
    keyCount = 0
    ReDim keys(0 To 200) ' Array más grande para evitar desbordamientos
    
    ' Limpiar el objeto JSON
    jsonObject = Trim(jsonObject)
    
    ' Buscar todas las claves en el objeto JSON
    pos = 1
    Do While pos <= Len(jsonObject) And pos > 0
        ' Buscar el patrón "clave":
        keyStart = InStr(pos, jsonObject, """")
        If keyStart = 0 Then Exit Do
        
        keyEnd = InStr(keyStart + 1, jsonObject, """")
        If keyEnd = 0 Then Exit Do
        
        key = Mid(jsonObject, keyStart + 1, keyEnd - keyStart - 1)
        
        ' Verificar que después de la clave hay ":"
        If keyEnd + 1 <= Len(jsonObject) Then
            char = Mid(jsonObject, keyEnd + 1, 1)
            If char = ":" Then
                ' Agregar clave al array si hay espacio
                If keyCount <= UBound(keys) Then
                    keys(keyCount) = key
                    keyCount = keyCount + 1
                End If
            End If
        End If
        
        pos = keyEnd + 1
    Loop
    
    ' Redimensionar array al tamaño correcto
    If keyCount > 0 Then
        ReDim Preserve keys(0 To keyCount - 1)
    Else
        ReDim keys(0 To 0)
        keys(0) = "Sin claves"
    End If
    
    ExtractKeysImproved = keys
    Exit Function
    
ErrorHandler:
    ' En caso de error, devolver array con mensaje de error
    ReDim keys(0 To 0)
    keys(0) = "Error: " & Err.Description
    ExtractKeysImproved = keys
    Debug.Print "ExtractKeysImproved - Error: " & Err.Description & " en posición " & pos
End Function

' Función auxiliar para extraer valores de un objeto JSON (versión mejorada)
Private Function ExtractValuesImproved(jsonObject As String) As String()
    On Error GoTo ErrorHandler
    
    Dim values() As String
    Dim valueCount As Integer
    Dim pos As Long
    Dim valueStart As Long
    Dim valueEnd As Long
    Dim value As String
    Dim char As String
    Dim inQuotes As Boolean
    Dim bracketCount As Integer
    
    ' Inicializar array con tamaño suficiente
    valueCount = 0
    ReDim values(0 To 200) ' Array más grande para evitar desbordamientos
    
    ' Limpiar el objeto JSON de espacios extra
    jsonObject = Trim(jsonObject)
    
    ' Buscar todos los valores después de ":"
    pos = 1
    Do While pos <= Len(jsonObject) And pos > 0
        ' Buscar ":"
        valueStart = InStr(pos, jsonObject, ":")
        If valueStart = 0 Then Exit Do
        
        pos = valueStart + 1
        
        ' Saltar espacios después de ":"
        Do While pos <= Len(jsonObject) And Mid(jsonObject, pos, 1) = " "
            pos = pos + 1
        Loop
        
        If pos > Len(jsonObject) Then Exit Do
        
        ' Extraer el valor basado en el primer carácter
        char = Mid(jsonObject, pos, 1)
        
        If char = """" Then
            ' Valor string - buscar la comilla de cierre
            valueStart = pos + 1
            valueEnd = InStr(valueStart, jsonObject, """")
            If valueEnd = 0 Then
                ' Si no encuentra comilla de cierre, tomar hasta el final
                value = Mid(jsonObject, valueStart)
                pos = Len(jsonObject) + 1
            Else
                value = Mid(jsonObject, valueStart, valueEnd - valueStart)
                pos = valueEnd + 1
            End If
        ElseIf char = "[" Then
            ' Valor array - buscar el ] correspondiente
            valueStart = pos
            pos = pos + 1
            bracketCount = 1
            Do While pos <= Len(jsonObject) And bracketCount > 0
                If Mid(jsonObject, pos, 1) = "[" Then
                    bracketCount = bracketCount + 1
                ElseIf Mid(jsonObject, pos, 1) = "]" Then
                    bracketCount = bracketCount - 1
                End If
                pos = pos + 1
            Loop
            value = Mid(jsonObject, valueStart, pos - valueStart)
        Else
            ' Valor numérico, null, true, false - buscar hasta coma o }
            valueStart = pos
            Do While pos <= Len(jsonObject)
                char = Mid(jsonObject, pos, 1)
                If char = "," Or char = "}" Then Exit Do
                pos = pos + 1
            Loop
            value = Mid(jsonObject, valueStart, pos - valueStart)
        End If
        
        ' Limpiar el valor
        value = Trim(value)
        
        ' Agregar al array si hay espacio
        If valueCount <= UBound(values) Then
            values(valueCount) = value
            valueCount = valueCount + 1
        End If
        
        ' Saltar coma y espacios
        Do While pos <= Len(jsonObject)
            char = Mid(jsonObject, pos, 1)
            If char <> "," And char <> " " Then Exit Do
            pos = pos + 1
        Loop
    Loop
    
    ' Redimensionar array al tamaño correcto
    If valueCount > 0 Then
        ReDim Preserve values(0 To valueCount - 1)
    Else
        ReDim values(0 To 0)
        values(0) = "Sin valores"
    End If
    
    ExtractValuesImproved = values
    Exit Function
    
ErrorHandler:
    ' En caso de error, devolver array con mensaje de error
    ReDim values(0 To 0)
    values(0) = "Error: " & Err.Description
    ExtractValuesImproved = values
    Debug.Print "ExtractValuesImproved - Error: " & Err.Description & " en posición " & pos
End Function

' Función auxiliar para dividir el array de objetos JSON de forma más robusta
Private Function SplitObjectsFromArray(dataArray As String) As String()
    On Error GoTo ErrorHandler
    
    Dim objects() As String
    Dim objectCount As Integer
    Dim pos As Long
    Dim objectStart As Long
    Dim objectEnd As Long
    Dim bracketCount As Integer
    Dim i As Long
    
    ' Inicializar array
    objectCount = 0
    ReDim objects(0 To 100) ' Array temporal con espacio suficiente
    
    ' Limpiar corchetes del array
    If Left(dataArray, 1) = "[" Then
        dataArray = Mid(dataArray, 2)
    End If
    If Right(dataArray, 1) = "]" Then
        dataArray = Left(dataArray, Len(dataArray) - 1)
    End If
    
    ' Buscar objetos individuales contando llaves
    pos = 1
    Do While pos <= Len(dataArray)
        ' Buscar el inicio de un objeto {
        objectStart = InStr(pos, dataArray, "{")
        If objectStart = 0 Then Exit Do
        
        ' Encontrar el final del objeto contando llaves
        bracketCount = 1
        i = objectStart + 1
        Do While i <= Len(dataArray) And bracketCount > 0
            If Mid(dataArray, i, 1) = "{" Then
                bracketCount = bracketCount + 1
            ElseIf Mid(dataArray, i, 1) = "}" Then
                bracketCount = bracketCount - 1
            End If
            i = i + 1
        Loop
        
        If bracketCount = 0 Then
            objectEnd = i - 1
            ' Agregar objeto al array
            If objectCount <= UBound(objects) Then
                objects(objectCount) = Mid(dataArray, objectStart, objectEnd - objectStart + 1)
                objectCount = objectCount + 1
            End If
            pos = objectEnd + 1
        Else
            Exit Do
        End If
    Loop
    
    ' Redimensionar array al tamaño correcto
    If objectCount > 0 Then
        ReDim Preserve objects(0 To objectCount - 1)
    Else
        ReDim objects(0 To 0)
        objects(0) = ""
    End If
    
    SplitObjectsFromArray = objects
    Exit Function
    
ErrorHandler:
    ReDim objects(0 To 0)
    objects(0) = ""
    SplitObjectsFromArray = objects
    Debug.Print "SplitObjectsFromArray - Error: " & Err.Description
End Function

' Función auxiliar para limpiar un objeto JSON
Private Function CleanJsonObject(jsonObject As String) As String
    On Error GoTo ErrorHandler
    
    ' Limpiar espacios
    jsonObject = Trim(jsonObject)
    
    ' Limpiar llaves del objeto
    If Left(jsonObject, 1) = "[" Then
        jsonObject = Mid(jsonObject, 2)
    End If
    If Left(jsonObject, 1) = "{" Then
        jsonObject = Mid(jsonObject, 2)
    End If
    If Right(jsonObject, 1) = "}" Then
        jsonObject = Left(jsonObject, Len(jsonObject) - 1)
    End If
    If Right(jsonObject, 1) = "]" Then
        jsonObject = Left(jsonObject, Len(jsonObject) - 1)
    End If
    
    CleanJsonObject = jsonObject
    Exit Function
    
ErrorHandler:
    CleanJsonObject = ""
    Debug.Print "CleanJsonObject - Error: " & Err.Description
End Function

' Función auxiliar para pintar tabla en Excel (versión mejorada sin ScriptControl)
Private Sub PaintTableInExcel(targetCell As Range, jsonData As String)
    On Error GoTo ErrorHandler
    
    Dim dataArray As String
    Dim objects() As String
    Dim keys() As String
    Dim values() As String
    Dim i As Integer
    Dim j As Integer
    Dim currentObject As String
    
    ' Parsear el JSON para extraer el array de data
    dataArray = ParseJsonDataImproved(jsonData)
    If dataArray = "" Then Exit Sub
    
    ' Dividir el array en objetos individuales con mejor lógica
    objects = SplitObjectsFromArray(dataArray)
    
    ' Verificar que tenemos al menos un objeto
    If UBound(objects) < 0 Then Exit Sub
    
    ' Obtener las claves del primer objeto para los encabezados
    currentObject = CleanJsonObject(objects(0))
    keys = ExtractKeysImproved(currentObject)
    
    ' Pintar encabezados UNA SOLA VEZ en la primera fila (3 filas debajo de la celda objetivo)
    For i = 0 To UBound(keys)
        targetCell.Offset(3, i).Value = keys(i)
    Next i
    
    ' Pintar datos en las siguientes filas (cada objeto es una fila de datos)
    For i = 0 To UBound(objects)
        ' Limpiar el objeto actual
        currentObject = CleanJsonObject(objects(i))
        
        ' Obtener valores del objeto actual
        values = ExtractValuesImproved(currentObject)
        
        ' Pintar SOLO LOS VALORES en la fila correspondiente (4 + i filas debajo de la celda objetivo)
        For j = 0 To UBound(values)
            If j <= UBound(keys) Then
                targetCell.Offset(4 + i, j).Value = values(j)
            End If
        Next j
    Next i
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "PaintTableInExcel - Error: " & Err.Description
End Sub

' Función auxiliar reutilizable para llamar a api_xls.f_pla_get_set_data_stock
Private Function ExecuteGetSetDataStock( _
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
    sqlQuery = "SELECT api_xls.f_pla_get_set_data_stock(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    
    On Error GoTo ErrorHandler
    
    ' Obtener conexión global (reutiliza la existente)
    Set conn = GetGlobalConnection()
    If conn Is Nothing Then
        ExecuteGetSetDataStock = "Error: No se pudo establecer conexión"
        Debug.Print "ExecuteGetSetDataStock - Error: No se pudo establecer conexión"
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
    
    ' Obtener resultado
    If Not rs.EOF Then
        ExecuteGetSetDataStock = rs.Fields(0).Value
    Else
        ExecuteGetSetDataStock = "No data found"
    End If
    
    ' Limpiar solo el recordset y comando (NO la conexión)
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    On Error GoTo 0
    
    Exit Function
    
ErrorHandler:
    ExecuteGetSetDataStock = "Error: " & Err.Description & " (Error #" & Err.Number & ")"
    Debug.Print "ExecuteGetSetDataStock - Error: " & Err.Description & " (Error #" & Err.Number & ")"
    ' Limpiar en caso de error (NO la conexión global)
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    On Error GoTo 0
End Function

' Función auxiliar para detectar si un valor parece ser numérico
Private Function IsNumericValue(value As String) As Boolean
    On Error GoTo ErrorHandler
    
    value = Trim(value)
    
    ' Casos especiales
    If value = "" Or LCase(value) = "null" Then
        IsNumericValue = False
        Exit Function
    End If
    
    ' Verificar si contiene solo dígitos, punto, coma y signos
    Dim i As Integer
    Dim char As String
    Dim hasDigit As Boolean
    Dim hasDecimal As Boolean
    
    hasDigit = False
    hasDecimal = False
    
    For i = 1 To Len(value)
        char = Mid(value, i, 1)
        
        If char >= "0" And char <= "9" Then
            hasDigit = True
        ElseIf char = "." Or char = "," Then
            If hasDecimal Then ' Ya tiene un separador decimal
                IsNumericValue = False
                Exit Function
            End If
            hasDecimal = True
        ElseIf char = "-" And i = 1 Then ' Signo negativo al inicio
            ' Permitir
        Else
            IsNumericValue = False
            Exit Function
        End If
    Next i
    
    IsNumericValue = hasDigit
    Exit Function
    
ErrorHandler:
    IsNumericValue = False
End Function

' Función auxiliar para convertir string a número si es posible
Private Function ConvertToNumberIfPossible(value As String) As Variant
    On Error GoTo ErrorHandler
    
    ' Limpiar espacios
    value = Trim(value)
    
    ' Si está vacío, devolver vacío
    If value = "" Then
        ConvertToNumberIfPossible = ""
        Exit Function
    End If
    
    ' Si es null, devolver vacío
    If LCase(value) = "null" Then
        ConvertToNumberIfPossible = ""
        Exit Function
    End If
    
    ' Reemplazar punto por coma para formato español
    Dim spanishValue As String
    spanishValue = Replace(value, ".", ",")
    
    ' Intentar convertir a número
    Dim numericValue As Double
    numericValue = CDbl(spanishValue)
    
    ' Si llegó aquí, es un número válido
    ConvertToNumberIfPossible = numericValue
    Exit Function
    
ErrorHandler:
    ' Si no se puede convertir, devolver como texto
    ConvertToNumberIfPossible = value
End Function

' =============================================================================
' FUNCIÓN AUXILIAR REUTILIZABLE PARA DEVOLVER ARRAY 2D
' =============================================================================

' Función auxiliar que contiene toda la lógica común para obtener datos y devolver Array 2D
Private Function ExecuteAndReturnArray2D( _
    ByVal p_unidad_operacional As Variant, _
    ByVal p_peticion As Variant, _
    ByVal p_producto_venta As Variant, _
    ByVal p_fecha_dato As Variant, _
    ByVal p_fch_inicial As Variant, _
    ByVal p_fch_final As Variant, _
    ByVal p_DiaVida_inicial As Variant, _
    ByVal p_DiaVida_final As Variant, _
    ByVal p_peso_inicial As Variant, _
    ByVal p_peso_final As Variant, _
    ByVal functionName As String) As Variant
    
    On Error GoTo ErrorHandler
    Dim jsonResult As String
    Dim dataArray As String
    Dim objects() As String
    Dim keys() As String
    Dim values() As String
    Dim i As Integer
    Dim j As Integer
    Dim currentObject As String
    Dim resultArray() As Variant
    
    ' Ejecutar la consulta a PostgreSQL
    jsonResult = ExecuteGetSetDataStock( _
        p_unidad_operacional, _
        p_peticion, _
        p_producto_venta, _
        p_fecha_dato, _
        p_fch_inicial, _
        p_fch_final, _
        p_DiaVida_inicial, _
        p_DiaVida_final, _
        p_peso_inicial, _
        p_peso_final)
    
    ' Si hay error, devolver el error
    If Left(jsonResult, 5) = "Error" Then
        ExecuteAndReturnArray2D = jsonResult
        Exit Function
    End If
    
    ' Parsear el JSON para extraer el array de data
    dataArray = ParseJsonDataImproved(jsonResult)
    If dataArray = "" Then
        ExecuteAndReturnArray2D = "Error: No se pudo extraer data del JSON"
        Exit Function
    End If
    
    ' Dividir el array en objetos individuales
    objects = SplitObjectsFromArray(dataArray)
    
    ' Verificar que tenemos al menos un objeto
    If UBound(objects) < 0 Then
        ExecuteAndReturnArray2D = "Error: No se encontraron objetos en el JSON"
        Exit Function
    End If
    
    ' Obtener las claves del primer objeto
    currentObject = CleanJsonObject(objects(0))
    keys = ExtractKeysImproved(currentObject)
    
    ' Crear array 2D: (filas = objetos + 1, columnas = claves)
    ReDim resultArray(1 To UBound(objects) + 2, 1 To UBound(keys) + 1)
    
    ' Llenar encabezados (fila 1)
    For i = 0 To UBound(keys)
        resultArray(1, i + 1) = keys(i)
    Next i
    
    ' Llenar datos (filas 2 en adelante)
    For i = 0 To UBound(objects)
        currentObject = CleanJsonObject(objects(i))
        values = ExtractValuesImproved(currentObject)
        
        For j = 0 To UBound(values)
            If j <= UBound(keys) Then
                ' CONVERSIÓN AUTOMÁTICA A NÚMERO
                If IsNumericValue(values(j)) Then
                    resultArray(i + 2, j + 1) = ConvertToNumberIfPossible(values(j))
                Else
                    resultArray(i + 2, j + 1) = values(j)
                End If
            End If
        Next j
    Next i
    
    ' Devolver el Array 2D
    ExecuteAndReturnArray2D = resultArray
    
    Exit Function
    
ErrorHandler:
    ExecuteAndReturnArray2D = "Error: " & Err.Description & " (Error #" & Err.Number & ")"
    Debug.Print "ExecuteAndReturnArray2D - Error: " & Err.Description & " (Error #" & Err.Number & ")"
End Function

' =============================================================================
' FUNCIONES PRINCIPALES PARA OBTENER Y DEVOLVER DATOS DE STOCK COMO ARRAY 2D
' =============================================================================

' Función GetSetStockAvi - Obtiene datos de stock y devuelve Array 2D
Public Function GetSetStockAvi(p_unidad_operacional As String, p_peticion As String, p_producto_venta As String, p_fecha_dato As Date, p_DiaVida_inicial As Integer, p_DiaVida_final As Integer, p_peso_inicial As Double, p_peso_final As Double) As Variant
    Dim fechaHoy As Date: fechaHoy = Date
    
    ' Usar la función auxiliar reutilizable con los parámetros mapeados
    GetSetStockAvi = ExecuteAndReturnArray2D( _
        p_unidad_operacional, _
        p_peticion, _
        p_producto_venta, _
        p_fecha_dato, _
        fechaHoy, _
        fechaHoy, _
        p_DiaVida_inicial, _
        p_DiaVida_final, _
        p_peso_inicial, _
        p_peso_final, _
        "GetSetStockAvi")
End Function

' Función GetSetEntradaAvi - Obtiene datos de entradas y devuelve Array 2D
Public Function GetSetEntradaAvi(p_unidad_operacional As String, p_peticion As String, p_producto_venta As String, p_fch_inicial As Date, p_fch_final As Date) As Variant
    Dim fechaHoy As Date: fechaHoy = Date
    
    ' Usar la función auxiliar reutilizable con los parámetros mapeados
    GetSetEntradaAvi = ExecuteAndReturnArray2D( _
        p_unidad_operacional, _
        p_peticion, _
        p_producto_venta, _
        fechaHoy, _
        p_fch_inicial, _
        p_fch_final, _
        0, _
        9999, _
        0, _
        99.999, _
        "GetSetEntradaAvi")
End Function

' Función GetSetSalidasAvi - Obtiene datos de salidas y devuelve Array 2D
Public Function GetSetSalidasAvi(p_unidad_operacional As String, p_peticion As String, p_producto_venta As String, p_fch_inicial As Date, p_fch_final As Date, p_DiaVida_inicial As Integer, p_DiaVida_final As Integer, p_peso_inicial As Double, p_peso_final As Double) As Variant
    Dim fechaHoy As Date: fechaHoy = Date
    
    ' Usar la función auxiliar reutilizable con los parámetros mapeados
    GetSetSalidasAvi = ExecuteAndReturnArray2D( _
        p_unidad_operacional, _
        p_peticion, _
        p_producto_venta, _
        fechaHoy, _
        p_fch_inicial, _
        p_fch_final, _
        p_DiaVida_inicial, _
        p_DiaVida_final, _
        p_peso_inicial, _
        p_peso_final, _
        "GetSetSalidasAvi")
End Function

' =============================================================================
' FUNCIONES DE UTILIDAD ADICIONALES
' =============================================================================

' Función para limpiar tablas pintadas (opcional)
Public Sub ClearPaintedTables(targetCell As Range, Optional numRows As Integer = 50, Optional numCols As Integer = 20)
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim j As Integer
    
    ' Limpiar el área de la tabla (3 filas debajo de la celda objetivo)
    For i = 3 To 3 + numRows
        For j = 0 To numCols - 1
            targetCell.Offset(i, j).Value = ""
        Next j
    Next i
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ClearPaintedTables - Error: " & Err.Description
End Sub

' Función para formatear tabla pintada (opcional)
Public Sub FormatPaintedTable(targetCell As Range, Optional numRows As Integer = 50, Optional numCols As Integer = 20)
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim j As Integer
    
    ' Formatear encabezados (fila 3)
    For j = 0 To numCols - 1
        With targetCell.Offset(3, j)
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
        End With
    Next j
    
    ' Aplicar bordes a toda la tabla
    For i = 3 To 3 + numRows
        For j = 0 To numCols - 1
            With targetCell.Offset(i, j).Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        Next j
    Next i
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "FormatPaintedTable - Error: " & Err.Description
End Sub
