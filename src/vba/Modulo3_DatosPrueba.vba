Sub InsertarDatosDePrueba()
    Dim wsF As Worksheet: Set wsF = ThisWorkbook.Sheets("Facturas")
    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Sheets("Datos")

    Dim cantidad As Variant
    cantidad = Application.InputBox("¿Cuántas facturas deseas insertar?", "Cantidad de Facturas", 10, Type:=1)
    If cantidad = False Then Exit Sub
    If Not IsNumeric(cantidad) Or cantidad < 1 Then
        MsgBox "Por favor, introduce un número válido mayor que cero.", vbExclamation
        Exit Sub
    End If

    Dim filaF As Long: filaF = wsF.Cells(wsF.Rows.count, "A").End(xlUp).row + 1
    Dim filaD As Long: filaD = wsD.Cells(wsD.Rows.count, "A").End(xlUp).row + 1

    ' Datos predefinidos válidos
    Dim tiposFactura As Variant: tiposFactura = Array("HONORARIO", "COMBUSTIBLE", "H&C")
    Dim nombres As Variant: nombres = Array("LUIS", "ANA", "CARLOS", "JUAN", "MARIA", "PEDRO", "ANDRES", "LAURA", "DANIEL", "SOFIA", _
                                             "JORGE", "CAMILA", "VICTOR", "ISABEL", "MARIO", "PATRICIA", "RICARDO", "NATALIA", "ALFREDO", "GABRIELA")
    Dim apellidos As Variant: apellidos = Array("PEREZ", "SOTO", "DIAZ", "BRAVO", "RUIZ", "LEON", "GOMEZ", "VARGAS", "PENA", "MARTINEZ", _
                                                "FERNANDEZ", "SALAZAR", "RAMIREZ", "CASTRO", "MORALES", "SANCHEZ", "GUTIERREZ", "DOMINGUEZ", "FUENTES", "LOPEZ")
    Dim aeronaves As Variant: aeronaves = Array("C172", "PA28", "DA40", "C150", "C152", "PIPER PA-18", "CESSNA 182", "MOONEY M20", "PIPER ARROW", "DIAMOND DA42", _
                                                "BEECHCRAFT BONANZA", "CESSNA 310", "PIPER SENECA", "DIAMOND DA20", "BEECH DUKE", "PIPER WARRIOR", "CIRRUS SR22", "PIPER ARCHER", "CUBCRAFTERS XCUB", "PIPER CHEROKEE")
    Dim metodosPago As Variant: metodosPago = Array("DIVISAS", "BOLIVARES")
    Dim codigosBanco As Variant: codigosBanco = Array("0102", "0104", "0105", "0108", "0114", "0115", "0116", "0128", "0134", "0137", "0138", "0146", "0151", "0156", "0163", "0166", "0168", "0171", "0172", "0173", "0174", "0175", "0177", "0191")

    Randomize

    Dim i As Long
    For i = 1 To cantidad
        Dim tipoFactura As String: tipoFactura = tiposFactura(Int(Rnd() * (UBound(tiposFactura) + 1)))
        Dim metodoPago As String: metodoPago = metodosPago(Int(Rnd() * (UBound(metodosPago) + 1)))
        Dim cedulaAlumno As String: cedulaAlumno = GenerarCedula()
        Dim fechaRecibo As Date: fechaRecibo = Now - Rnd() * 30
        Dim fechaVuelo As Date: fechaVuelo = Date - Int(Rnd() * 10)
        
        ' Crear nombres aleatorios
        Dim nombreInstructor As String: nombreInstructor = nombres(Int(Rnd() * (UBound(nombres) + 1)))
        Dim apellidoInstructor As String: apellidoInstructor = apellidos(Int(Rnd() * (UBound(apellidos) + 1)))
        Dim nombreAlumno As String: nombreAlumno = nombres(Int(Rnd() * (UBound(nombres) + 1)))
        Dim apellidoAlumno As String: apellidoAlumno = apellidos(Int(Rnd() * (UBound(apellidos) + 1)))
        Dim nombreDespachador As String: nombreDespachador = nombres(Int(Rnd() * (UBound(nombres) + 1))) & " " & apellidos(Int(Rnd() * (UBound(apellidos) + 1)))

        Dim nombreInstructorCompleto As String: nombreInstructorCompleto = nombreInstructor & " " & apellidoInstructor
        Dim nombreAlumnoCompleto As String: nombreAlumnoCompleto = nombreAlumno & " " & apellidoAlumno

        ' === Insertar en FACTURAS ===
        wsF.Cells(filaF, 1).Value = GenerarIDUnico()
        wsF.Cells(filaF, 2).Value = Format(fechaRecibo, "dd/mm/yyyy hh:mm")
        wsF.Cells(filaF, 3).Value = Format(fechaVuelo, "dd/mm/yyyy")
        wsF.Cells(filaF, 4).Value = nombreDespachador

        If tipoFactura <> "COMBUSTIBLE" Then
            wsF.Cells(filaF, 5).Value = nombreInstructorCompleto
        Else
            wsF.Cells(filaF, 5).Value = ""
        End If

        wsF.Cells(filaF, 6).Value = nombreAlumnoCompleto
        wsF.Cells(filaF, 7).Value = cedulaAlumno

        If tipoFactura <> "HONORARIO" Then
            wsF.Cells(filaF, 8).Value = aeronaves(Int(Rnd() * (UBound(aeronaves) + 1)))
            wsF.Cells(filaF, 9).Value = Int(10 + Rnd() * 40)
        Else
            wsF.Cells(filaF, 8).Value = ""
            wsF.Cells(filaF, 9).Value = ""
        End If

        Dim monto As Long: monto = CLng((20 + Rnd() * 200) * 1000)
        wsF.Cells(filaF, 10).Value = metodoPago
        wsF.Cells(filaF, 11).Value = monto

        If metodoPago = "BOLIVARES" Then
            Dim bancoCodigo As String: bancoCodigo = codigosBanco(Int(Rnd() * (UBound(codigosBanco) + 1)))
            Dim cedulaDepositante As String: cedulaDepositante = cedulaAlumno
            Dim numeroOperacion As String: numeroOperacion = Format(Int(Rnd() * 10 ^ 10), "0000000000")
            Dim telefonoOrigen As String: telefonoOrigen = "04" & CStr(Int(1 + Rnd() * 3)) & Format(Int(Rnd() * 10000000), "0000000")

            wsF.Cells(filaF, 12).Value = bancoCodigo
            wsF.Cells(filaF, 13).Value = cedulaDepositante
            wsF.Cells(filaF, 14).Value = numeroOperacion
            wsF.Cells(filaF, 15).Value = telefonoOrigen
        Else
            wsF.Cells(filaF, 12).Value = ""
            wsF.Cells(filaF, 13).Value = ""
            wsF.Cells(filaF, 14).Value = ""
            wsF.Cells(filaF, 15).Value = ""
        End If

        wsF.Cells(filaF, 16).Value = tipoFactura
        wsF.Cells(filaF, 17).Value = "OBSERVACIÓN DE PRUEBA " & filaF - 1

        ' === Insertar en DATOS coherentes ===

        ' Instructor
        If tipoFactura <> "COMBUSTIBLE" Then
            If Not ExisteEnDatos(wsD, nombreInstructor, apellidoInstructor, "NO APLICA", "INSTRUCTOR") Then
                wsD.Cells(filaD, 1).Value = GenerarIDUnico()
                wsD.Cells(filaD, 2).Value = "INSTRUCTOR"
                wsD.Cells(filaD, 3).Value = nombreInstructor
                wsD.Cells(filaD, 4).Value = apellidoInstructor
                wsD.Cells(filaD, 5).Value = "NO APLICA"
                filaD = filaD + 1
            End If
        End If

        ' Alumno
        If Not ExisteEnDatos(wsD, nombreAlumno, apellidoAlumno, cedulaAlumno, "ALUMNO") Then
            wsD.Cells(filaD, 1).Value = GenerarIDUnico()
            wsD.Cells(filaD, 2).Value = "ALUMNO"
            wsD.Cells(filaD, 3).Value = nombreAlumno
            wsD.Cells(filaD, 4).Value = apellidoAlumno
            wsD.Cells(filaD, 5).Value = cedulaAlumno
            filaD = filaD + 1
        End If

        filaF = filaF + 1
    Next i

    MsgBox cantidad & " facturas y datos de prueba insertados correctamente.", vbInformation
End Sub

' Función para verificar duplicados en Datos
Private Function ExisteEnDatos(ws As Worksheet, nombre As String, apellido As String, cedula As String, cargo As String) As Boolean
    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    Dim i As Long
    For i = 2 To ultimaFila
        If _
            Trim(UCase(ws.Cells(i, 2).Value)) = Trim(UCase(cargo)) And _
            Trim(UCase(ws.Cells(i, 3).Value)) = Trim(UCase(nombre)) And _
            Trim(UCase(ws.Cells(i, 4).Value)) = Trim(UCase(apellido)) And _
            Trim(UCase(ws.Cells(i, 5).Value)) = Trim(UCase(cedula)) Then
            
            ExisteEnDatos = True
            Exit Function
        End If
    Next i
    
    ExisteEnDatos = False
End Function

Private Function GenerarCedula() As String
    Dim prefijos As Variant: prefijos = Array("V", "E")
    Dim prefijo As String: prefijo = prefijos(Int(Rnd() * 2))
    Dim num As Long: num = CLng(1000000 + Rnd() * 89999999)
    GenerarCedula = prefijo & num
End Function

Function GenerarIDUnico() As String
    Dim tiempo As String
    Dim aleatorio As String
    Dim letras As String
    Dim letra1 As String, letra2 As String
    Dim idGenerado As String
    Dim existe As Boolean
    Dim wsDatos As Worksheet, wsFacturas As Worksheet
    Dim rngDatos As Range, rngFacturas As Range
    Dim celda As Range

    letras = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Set wsDatos = ThisWorkbook.Sheets("Datos")
    Set wsFacturas = ThisWorkbook.Sheets("Facturas")
    
    Do
        ' Generar ID
        Randomize
        letra1 = Mid(letras, Int(26 * Rnd + 1), 1)
        letra2 = Mid(letras, Int(26 * Rnd + 1), 1)
        tiempo = Format(Now, "yyyymmdd-hhnnss")
        aleatorio = Format(Int((9999 - 1000 + 1) * Rnd + 1000), "0000")
        idGenerado = letra1 & letra2 & "-" & tiempo & "-" & aleatorio
        
        ' Comprobar existencia en hoja "Datos"
        Set rngDatos = wsDatos.Range("A:A")
        Set celda = rngDatos.Find(What:=idGenerado, LookIn:=xlValues, LookAt:=xlWhole)
        existe = Not celda Is Nothing
        
        ' Si no está en "Datos", comprobar en "Facturas"
        If Not existe Then
            Set rngFacturas = wsFacturas.Range("A:A")
            Set celda = rngFacturas.Find(What:=idGenerado, LookIn:=xlValues, LookAt:=xlWhole)
            existe = Not celda Is Nothing
        End If
    Loop While existe

    GenerarIDUnico = idGenerado
End Function

Sub ExportarCSV()
    Dim ruta As String
    Dim fechaNombre As String
    
    ' Verificar si el libro está guardado
    If ThisWorkbook.Path = "" Then
        MsgBox "Guarda el libro antes de exportar.", vbExclamation
        Exit Sub
    End If
    
    fechaNombre = Format(Date, "yyyymmdd")
    ruta = ThisWorkbook.Path & "\"
    
    On Error Resume Next ' Para manejar errores
    
    ' Exportar hoja "Facturas"
    If SheetExists("Facturas") Then
        Sheets("Facturas").Copy
        With ActiveWorkbook
            .SaveAs Filename:=ruta & "Facturas_" & fechaNombre & ".csv", FileFormat:=6 ' 6 = xlCSV
            .Close False
        End With
    Else
        MsgBox "La hoja 'Facturas' no existe.", vbExclamation
    End If
    
    ' Exportar hoja "Datos"
    If SheetExists("Datos") Then
        Sheets("Datos").Copy
        With ActiveWorkbook
            .SaveAs Filename:=ruta & "Datos_" & fechaNombre & ".csv", FileFormat:=6
            .Close False
        End With
    Else
        MsgBox "La hoja 'Datos' no existe.", vbExclamation
    End If
    
    On Error GoTo 0 ' Restaura el manejo de errores
    
    MsgBox "Proceso de exportación completado.", vbInformation
End Sub

' Función auxiliar para verificar si existe una hoja
Function SheetExists(sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = (Sheets(sheetName).Name <> "")
    On Error GoTo 0
End Function

Sub ImportarFacturasDesdeCSV()
    Dim ruta As String
    Dim archivo As String
    Dim ws As Worksheet

    Set ws = Sheets("Facturas")
    ruta = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Selecciona archivo de Facturas")

    If ruta = "False" Then Exit Sub

    With ws
        .Cells.ClearContents
        With .QueryTables.Add(Connection:="TEXT;" & ruta, Destination:=.Range("A1"))
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            .Refresh BackgroundQuery:=False
        End With
    End With

    MsgBox "Datos importados a 'Facturas'.", vbInformation
End Sub

Sub ImportarDatosDesdeCSV()
    Dim ruta As String
    Dim archivo As String
    Dim ws As Worksheet

    Set ws = Sheets("Datos")
    ruta = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Selecciona archivo de Datos")

    If ruta = "False" Then Exit Sub

    With ws
        .Cells.ClearContents
        With .QueryTables.Add(Connection:="TEXT;" & ruta, Destination:=.Range("A1"))
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            .Refresh BackgroundQuery:=False
        End With
    End With

    MsgBox "Datos importados a 'Datos'.", vbInformation
End Sub

-------------------------------------------------------------------------------