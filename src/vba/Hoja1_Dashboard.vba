Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    Application.EnableEvents = False
    
    ' Verificar si los cambios ocurren en las celdas de selección de gráficos (E3, E4, E5)
    If Not Intersect(Target, Me.Range("E3:E5")) Is Nothing Then
        ' Validar que las celdas contengan solo valores permitidos o estén vacías
        If IsValidChartSelection(Target.Value) Or Target.Value = "" Then
            Call MostrarOcultarGraficos
        Else
            MsgBox "Selección no válida. Por favor elija una opción de la lista desplegable.", vbExclamation, "Error en selección"
            Application.Undo
        End If
    End If
    
    ' Verificar si el cambio fue en la celda B6 (tasa Bs/USD)
    If Not Intersect(Target, Me.Range("B6")) Is Nothing Then
        ' Validar que sea un número válido
        If IsNumeric(Target.Value) And Target.Value > 0 Then
            'Call ActualizarDatos
        Else
            MsgBox "Por favor ingrese un valor numérico válido mayor que cero para la tasa.", vbExclamation, "Error en tasa"
            Application.Undo
        End If
    End If
    
CleanExit:
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Se produjo un error: " & Err.Description, vbCritical, "Error"
    Resume CleanExit
End Sub

' Función auxiliar para validar selecciones de gráficos
Private Function IsValidChartSelection(ByVal selectionValue As Variant) As Boolean
    Dim validValues As Variant
    validValues = Array("RECAUDACIÓN EN USD", "COMBUSTIBLE DESPACHADO", _
                       "INSTRUCTORES", "AERONAVES MÁS TRIPULADAS", _
                       "DISTRIBUCIÓN TIPO DE FACTURA", "DESPACHADOR")
    
    Dim i As Integer
    For i = LBound(validValues) To UBound(validValues)
        If UCase(selectionValue) = validValues(i) Then
            IsValidChartSelection = True
            Exit Function
        End If
    Next i
    
    IsValidChartSelection = False
End Function

' Función para mostrar/ocultar gráficos según selección en el panel
Private Sub MostrarOcultarGraficos()
    Dim wsDash As Worksheet
    Dim chartObj As ChartObject
    Dim graficoVisible As Boolean
    Dim tipoGrafico As String
    Dim i As Integer
    
    ' Configuración inicial
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Application.ScreenUpdating = False
    
    ' Ocultar todos los gráficos primero
    For Each chartObj In wsDash.ChartObjects
        chartObj.visible = False
    Next chartObj
    
    ' --- Controlar gráfico de barras verticales ---
    tipoGrafico = wsDash.Range("E3").Value
    graficoVisible = (tipoGrafico <> "")
    
    For Each chartObj In wsDash.ChartObjects
        With chartObj.chart
            ' Gráfico de barras verticales (puede ser Recaudación o Combustible)
            If .ChartType = xlColumnClustered Then
                If .ChartTitle.Text Like "*Recaudación*" And tipoGrafico Like "*RECAUDACIÓN*" Then
                    chartObj.visible = graficoVisible
                ElseIf .ChartTitle.Text Like "*Combustible*" And tipoGrafico Like "*COMBUSTIBLE*" Then
                    chartObj.visible = graficoVisible
                End If
            End If
        End With
    Next chartObj
    
    ' --- Controlar gráfico de barras horizontales ---
    tipoGrafico = wsDash.Range("E4").Value
    graficoVisible = (tipoGrafico <> "")
    
    For Each chartObj In wsDash.ChartObjects
        With chartObj.chart
            ' Gráfico de barras horizontales (Instructores o Aeronaves)
            If .ChartType = xlBarClustered Then
                If .ChartTitle.Text Like "*Instructores*" And tipoGrafico Like "*INSTRUCTORES*" Then
                    chartObj.visible = graficoVisible
                ElseIf .ChartTitle.Text Like "*Aeronaves*" And tipoGrafico Like "*AERONAVES*" Then
                    chartObj.visible = graficoVisible
                End If
            End If
        End With
    Next chartObj
    
    ' --- Controlar gráfico circular ---
    tipoGrafico = wsDash.Range("E5").Value
    graficoVisible = (tipoGrafico <> "")
    
    For Each chartObj In wsDash.ChartObjects
        With chartObj.chart
            ' Gráfico circular (Tipo de Factura o Despachador)
            If .ChartType = xlPie Then
                If .ChartTitle.Text Like "*Factura*" And tipoGrafico Like "*FACTURA*" Then
                    chartObj.visible = graficoVisible
                ElseIf .ChartTitle.Text Like "*Despachador*" And tipoGrafico Like "*DESPACHADOR*" Then
                    chartObj.visible = graficoVisible
                End If
            End If
        End With
    Next chartObj
    
    Application.ScreenUpdating = True
End Sub

Private Sub btnSincronizarFacturas_Click()
    Dim folderPath As String
    Dim fso As Object, folder As Object, file As Object
    Dim wb As Workbook
    Dim clave As String: clave = "clave123"
    Dim dataGlobal As Collection
    Dim arrCabecera As Variant
    Dim dicUnicos As Object
    Dim archivosProcesados As Long
    Dim totalRegistros As Long
    Dim librosParaActualizar As Collection
    Dim dicIDsGlobal As Object

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    'Hacer respaldo antes de sincronizar
    Call RespaldarAntesDeSincronizar

    Set dicUnicos = CreateObject("Scripting.Dictionary")
    Set dicIDsGlobal = CreateObject("Scripting.Dictionary")
    Set dataGlobal = New Collection
    Set librosParaActualizar = New Collection

    folderPath = ThisWorkbook.Path & "\"

    arrCabecera = Array("ID", "FECHA DE RECIBO", "FECHA DEL VUELO", "DESPACHADOR", "INSTRUCTOR", "ALUMNO", _
                        "CEDULA DEL ALUMNO", "AERONAVE", "CANTIDAD DE COMBUSTIBLE DEL AERONAVE", "TIPO DE PAGO", _
                        "PAGO", "BANCO", "CEDULA DEL DEPOSITANTE", "NUM DE OPERACIÓN", "ORIGEN (NUM. TELEFÓNICO)", _
                        "TIPO DE FACTURA", "OBSERVACIONES")

    ' Procesar archivo actual
    Call AgregarDatosUnicos(ThisWorkbook, "Facturas", clave, dataGlobal, dicUnicos, arrCabecera, dicIDsGlobal)
    archivosProcesados = 1

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    ' Procesar los demás archivos
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "xlsm" And file.Path <> ThisWorkbook.FullName Then
            On Error Resume Next
            Set wb = Workbooks.Open(file.Path, Password:=clave)
            On Error GoTo 0

            If Not wb Is Nothing Then
                Call AgregarDatosUnicos(wb, "Facturas", clave, dataGlobal, dicUnicos, arrCabecera, dicIDsGlobal)
                librosParaActualizar.Add wb
                archivosProcesados = archivosProcesados + 1
            Else
                Call RegistrarError(file.Name, "No se pudo abrir el archivo.", "ERROR")
            End If
        End If
    Next file

    ' Escribir datos unificados en todas las hojas
    Call EscribirDatosEnFacturas(ThisWorkbook, "Facturas", clave, dataGlobal, arrCabecera)

    For Each wb In librosParaActualizar
        On Error Resume Next
        Call EscribirDatosEnFacturas(wb, "Facturas", clave, dataGlobal, arrCabecera)
        wb.Close SaveChanges:=True
        On Error GoTo 0
    Next wb

    totalRegistros = dataGlobal.count

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    'Agregar nuevos usuarios a la hoja Datos
    Call AgregarNuevosDatosDesdeFacturas

    ' Mensaje final y log de éxito
    Dim mensaje As String
    mensaje = "Sincronización completada correctamente. " & totalRegistros & " registros únicos desde " & archivosProcesados & " archivos."
    MsgBox mensaje, vbInformation
    Call RegistrarError(ThisWorkbook.Name, mensaje, "ÉXITO")
End Sub

Private Sub btnRestaurarFacturas_Click()
    Dim wsR1 As Worksheet, wsR2 As Worksheet
    Dim wsFacturas As Worksheet, wsDatos As Worksheet
    Dim lastRowR1 As Long, lastRowR2 As Long
    Dim lastRowFacturas As Long, lastRowDatos As Long
    Dim continuar As VbMsgBoxResult

    ' Establecer referencias a las hojas
    Set wsR1 = ThisWorkbook.Sheets("R1")
    Set wsR2 = ThisWorkbook.Sheets("R2")
    Set wsFacturas = ThisWorkbook.Sheets("Facturas")
    Set wsDatos = ThisWorkbook.Sheets("Datos")

    ' Verificar si R1 o R2 tienen solo una fila de datos (2 filas en total)
    lastRowR1 = wsR1.Cells(wsR1.Rows.count, "A").End(xlUp).row
    lastRowR2 = wsR2.Cells(wsR2.Rows.count, "A").End(xlUp).row

    If lastRowR1 = 2 Or lastRowR2 = 2 Then
        continuar = MsgBox("Las hojas R1 y/o R2 solo tienen una fila de datos. ¿Deseas continuar con la restauración?", vbYesNo + vbQuestion, "Confirmar restauración")
        If continuar = vbNo Then
            MsgBox "Restauración cancelada.", vbExclamation
            Exit Sub
        End If
    End If

    ' ---------------------
    ' Procesar hoja Facturas
    ' ---------------------
    If lastRowR1 > 1 Then
        ' Borrar datos antiguos (excepto encabezados)
        lastRowFacturas = wsFacturas.Cells(wsFacturas.Rows.count, "A").End(xlUp).row
        If lastRowFacturas > 1 Then
            wsFacturas.Range("A2:Q" & lastRowFacturas).ClearContents
        End If
        
        ' Copiar datos desde R1
        wsR1.Range("A2:Q" & lastRowR1).Copy
        wsFacturas.Range("A2").PasteSpecial xlPasteValues
    End If

    ' ---------------------
    ' Procesar hoja Datos
    ' ---------------------
    If lastRowR2 > 1 Then
        lastRowDatos = wsDatos.Cells(wsDatos.Rows.count, "A").End(xlUp).row
        If lastRowDatos > 1 Then
            wsDatos.Range("A2:E" & lastRowDatos).ClearContents
        End If
        
        wsR2.Range("A2:E" & lastRowR2).Copy
        wsDatos.Range("A2").PasteSpecial xlPasteValues
    End If

    Application.CutCopyMode = False
    MsgBox "Datos restaurados correctamente.", vbInformation
End Sub

Private Sub AgregarDatosUnicos(wb As Workbook, hojaNombre As String, clave As String, _
                                dataGlobal As Collection, dicUnicos As Object, headers As Variant, _
                                dicIDs As Object)
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim datosFila As Variant
    Dim fila() As Variant
    Dim IDactual As String, claveDatos As String
    Dim fechaRecibo As String
    Dim lastRow As Long
    Dim dataRange As Range
    Dim estabaProtegida As Boolean

    On Error Resume Next
    Set ws = wb.Sheets(hojaNombre)
    On Error GoTo 0
    If ws Is Nothing Then
        Call RegistrarError(wb.Name, "No se encontró la hoja '" & hojaNombre & "'.")
        Exit Sub
    End If

    estabaProtegida = ws.ProtectContents
    If estabaProtegida Then On Error Resume Next: ws.Unprotect clave: On Error GoTo 0

    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow < 2 Then GoTo Salir ' No hay datos

    Set dataRange = ws.Range("A2:Q" & lastRow)

    For i = 1 To dataRange.Rows.count
        datosFila = Application.Index(dataRange.Value, i, 0)
        If IsArray(datosFila) Then
            IDactual = Trim(CStr(datosFila(1)))
            fechaRecibo = Trim(CStr(datosFila(2)))

            claveDatos = ""
            For j = 2 To 17
                claveDatos = claveDatos & "|" & Trim(CStr(datosFila(j)))
            Next j

            If dicIDs.Exists(IDactual) Then
                If dicIDs(IDactual) <> fechaRecibo Then
                    ' Mismo ID, pero fecha distinta => nuevo ID
                    ReDim fila(1 To 17)
                    fila(1) = GenerarIDUnico()
                    For j = 2 To 17
                        fila(j) = datosFila(j)
                    Next j
                    If Not dicUnicos.Exists(claveDatos) Then
                        dicUnicos.Add claveDatos, True
                        dataGlobal.Add fila
                    End If
                End If
            Else
                ' Nuevo ID
                dicIDs.Add IDactual, fechaRecibo
                ReDim fila(1 To 17)
                For j = 1 To 17
                    fila(j) = datosFila(j)
                Next j
                If Not dicUnicos.Exists(claveDatos) Then
                    dicUnicos.Add claveDatos, True
                    dataGlobal.Add fila
                End If
            End If
        End If
    Next i

Salir:
    If estabaProtegida Then ws.Protect clave
End Sub

Private Sub EscribirDatosEnFacturas(wb As Workbook, hojaNombre As String, clave As String, _
                                    dataGlobal As Collection, headers As Variant)
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim lastRow As Long
    Dim fila() As Variant
    Dim dicLocal As Object
    Dim estabaProtegida As Boolean
    Dim filaEscritura As Long
    Dim idExistente As String

    Set dicLocal = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set ws = wb.Sheets(hojaNombre)
    On Error GoTo 0
    If ws Is Nothing Then
        Call RegistrarError(wb.Name, "No se encontró la hoja '" & hojaNombre & "' al escribir.")
        Exit Sub
    End If

    ' Desproteger hoja si estaba protegida
    estabaProtegida = ws.ProtectContents
    If estabaProtegida Then On Error Resume Next: ws.Unprotect clave: On Error GoTo 0

    ' Registrar todos los IDs existentes
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow >= 2 Then
        For i = 2 To lastRow
            idExistente = Trim(CStr(ws.Cells(i, 1).Value))
            If Len(idExistente) > 0 Then
                dicLocal(idExistente) = True
            End If
        Next i
    End If

    ' Buscar la siguiente fila disponible
    filaEscritura = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1

    ' Escribir datos no duplicados (basado en ID)
    For i = 1 To dataGlobal.count
        fila = dataGlobal(i)
        If Not dicLocal.Exists(fila(1)) Then
            ws.Range("A" & filaEscritura & ":Q" & filaEscritura).Value = fila
            dicLocal.Add fila(1), True
            filaEscritura = filaEscritura + 1
        End If
    Next i

    ' Volver a proteger la hoja si estaba protegida
    If estabaProtegida Then ws.Protect clave
End Sub

Private Sub RespaldarAntesDeSincronizar()
    On Error Resume Next
    Dim wsOrigen As Worksheet, wsDestino As Worksheet

    ' Copiar Facturas a R1
    Set wsOrigen = ThisWorkbook.Sheets("Facturas")
    Set wsDestino = ThisWorkbook.Sheets("R1")
    If Not wsOrigen Is Nothing And Not wsDestino Is Nothing Then
        wsDestino.Cells.Clear
        wsOrigen.Cells.Copy Destination:=wsDestino.Cells(1, 1)
    End If

    ' Copiar Datos a R2
    Set wsOrigen = ThisWorkbook.Sheets("Datos")
    Set wsDestino = ThisWorkbook.Sheets("R2")
    If Not wsOrigen Is Nothing And Not wsDestino Is Nothing Then
        wsDestino.Cells.Clear
        wsOrigen.Cells.Copy Destination:=wsDestino.Cells(1, 1)
    End If
End Sub

Private Sub AgregarNuevosDatosDesdeFacturas()
    Dim wsFacturas As Worksheet, wsDatos As Worksheet
    Dim ultimaFilaFacturas As Long, ultimaFilaDatos As Long
    Dim fila As Long
    Dim nombreCompleto As Variant
    Dim nombre As String, apellido As String, cedula As String, cargo As String
    Dim existe As Boolean
    Dim claveUnica As String
    Dim dicExistentes As Object

    Set dicExistentes = CreateObject("Scripting.Dictionary")
    Set wsFacturas = ThisWorkbook.Sheets("Facturas")
    Set wsDatos = ThisWorkbook.Sheets("Datos")

    ultimaFilaDatos = wsDatos.Cells(wsDatos.Rows.count, 1).End(xlUp).row
    ultimaFilaFacturas = wsFacturas.Cells(wsFacturas.Rows.count, 1).End(xlUp).row

    ' Guardar claves únicas de los existentes en "Datos"
    For fila = 2 To ultimaFilaDatos
        nombre = Trim(UCase(wsDatos.Cells(fila, 3).Value))
        apellido = Trim(UCase(wsDatos.Cells(fila, 4).Value))
        cedula = Trim(UCase(wsDatos.Cells(fila, 5).Value))
        claveUnica = nombre & "|" & apellido & "|" & cedula
        dicExistentes(claveUnica) = True
    Next fila

    ' Revisar instructores
    For fila = 2 To ultimaFilaFacturas
        ' Instructores
        nombreCompleto = Split(Trim(wsFacturas.Cells(fila, 5).Value))
        If UBound(nombreCompleto) >= 1 Then
            nombre = UCase(nombreCompleto(0))
            apellido = UCase(nombreCompleto(1))
            cedula = "NO APLICA"
            claveUnica = nombre & "|" & apellido & "|" & cedula
            If Not dicExistentes.Exists(claveUnica) Then
                ultimaFilaDatos = ultimaFilaDatos + 1
                wsDatos.Cells(ultimaFilaDatos, 1).Value = GenerarIDUnico()
                wsDatos.Cells(ultimaFilaDatos, 2).Value = "INSTRUCTOR"
                wsDatos.Cells(ultimaFilaDatos, 3).Value = nombre
                wsDatos.Cells(ultimaFilaDatos, 4).Value = apellido
                wsDatos.Cells(ultimaFilaDatos, 5).Value = cedula
                dicExistentes(claveUnica) = True
            End If
        End If

        ' Alumnos
        nombreCompleto = Split(Trim(wsFacturas.Cells(fila, 6).Value))
        If UBound(nombreCompleto) >= 1 Then
            nombre = UCase(nombreCompleto(0))
            apellido = UCase(nombreCompleto(1))
            cedula = Trim(wsFacturas.Cells(fila, 7).Value)
            claveUnica = nombre & "|" & apellido & "|" & UCase(cedula)
            If Not dicExistentes.Exists(claveUnica) Then
                ultimaFilaDatos = ultimaFilaDatos + 1
                wsDatos.Cells(ultimaFilaDatos, 1).Value = GenerarIDUnico()
                wsDatos.Cells(ultimaFilaDatos, 2).Value = "ALUMNO"
                wsDatos.Cells(ultimaFilaDatos, 3).Value = nombre
                wsDatos.Cells(ultimaFilaDatos, 4).Value = apellido
                wsDatos.Cells(ultimaFilaDatos, 5).Value = cedula
                dicExistentes(claveUnica) = True
            End If
        End If
    Next fila
End Sub

Private Function SliceArray(arr As Variant, startIdx As Long) As Variant
    Dim result() As Variant
    Dim i As Long
    ReDim result(0 To UBound(arr) - startIdx)
    For i = startIdx To UBound(arr)
        result(i - startIdx) = arr(i)
    Next i
    SliceArray = result
End Function

Private Sub RegistrarError(nombreArchivo As String, descripcion As String, Optional tipo As String = "ERROR")
    Dim logSheet As Worksheet
    On Error Resume Next
    Set logSheet = ThisWorkbook.Sheets("Log")
    If logSheet Is Nothing Then
        Set logSheet = ThisWorkbook.Sheets.Add
        logSheet.Name = "Log"
        logSheet.Range("A1:D1").Value = Array("Fecha", "Archivo", "Tipo", "Descripción")
    End If
    On Error GoTo 0

    With logSheet
        Dim fila As Long
        fila = .Cells(.Rows.count, 1).End(xlUp).row + 1
        .Cells(fila, 1).Value = Now
        .Cells(fila, 2).Value = nombreArchivo
        .Cells(fila, 3).Value = tipo
        .Cells(fila, 4).Value = descripcion
    End With
End Sub


+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Workbook_Open       |Runs when the Excel Workbook is opened       |
|AutoExec  |Workbook_BeforeClose|Runs when the Excel Workbook is closed       |
|AutoExec  |btnGuardar_Click    |Runs when the file is opened and ActiveX     |
|          |                    |objects trigger events                       |
|AutoExec  |cbxBanco_Change     |Runs when the file is opened and ActiveX     |
|          |                    |objects trigger events                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Create              |May execute file or a system command through |
|          |                    |WMI                                          |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|MkDir               |May create a directory                       |
|Suspicious|CreateObject        |May create an OLE object                     |
|Suspicious|ExecuteExcel4Macro  |May run an Excel 4 Macro (aka XLM/XLF) from  |
|          |                    |VBA                                          |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Base64    |%@                  |JUAN                                         |
|String    |                    |                                             |
|Base64    |<C@                 |PENA                                         |
|String    |                    |                                             |
|Base64    |N*h                 |Tipo                                         |
|String    |                    |                                             |
|Suspicious|VBA Stomping        |VBA Stomping was detected: the VBA source    |
|          |                    |code and P-code are different, this may have |
|          |                    |been used to hide malicious code             |
+----------+--------------------+---------------------------------------------+
VBA Stomping detection is experimental: please report any false positive/negative at https://github.com/decalage2/oletools/issues