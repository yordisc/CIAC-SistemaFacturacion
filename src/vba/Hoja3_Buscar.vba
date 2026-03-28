Private Sub Worksheet_Activate()
    Dim wsFacturas As Worksheet
    Dim hojaExtra As Worksheet
    Dim celda As Range
    Dim listaVisible() As String
    Dim listaAutoCompletar() As String
    Dim i As Long, j As Long
    
    ' Referencias
    Set wsFacturas = ThisWorkbook.Sheets("Facturas")
    On Error Resume Next
    Set hojaExtra = ThisWorkbook.Sheets("Extras")
    On Error GoTo 0
    
    If hojaExtra Is Nothing Then
        MsgBox "La hoja 'Extras' no existe.", vbCritical
        Exit Sub
    End If
    
    ' Limpiar ComboBoxes
    With Me.cbxFiltroDato
        .Clear
        .Style = fmStyleDropDownList
    End With
    With Me.cbxBuscarDato
        .Clear
        .Value = ""
        .MatchEntry = fmMatchEntryComplete
        .Style = fmStyleDropDownCombo
    End With
    
    ' Limpiar tabla Buscar (sin tocar encabezados de la fila 2)
    Me.Range("A3:P1048576").ClearContents
    
    ' Cargar encabezados de Facturas a cbxFiltroDato
    Dim excluidos As Variant
    excluidos = Array("ID", "CANTIDAD DE COMBUSTIBLE DEL AERONAVE", "OBSERVACIONES", "NUM DE OPERACIÓN", "PAGO")
    Dim ex
    Dim encabezado As String
    Dim agregar As Boolean
    Dim col As Integer
    
    For col = 1 To 17 ' Columnas A a Q en Facturas
        encabezado = wsFacturas.Cells(1, col).Value
        agregar = True
        For Each ex In excluidos
            If encabezado = ex Then
                agregar = False
                Exit For
            End If
        Next ex
        If agregar Then Me.cbxFiltroDato.AddItem encabezado
    Next col
    
    ' Cargar lista para búsqueda si cbxFiltroDato está vacío
    If Me.cbxFiltroDato.Value = "" Then
        Dim cntVisible As Long, cntAutoCompletar As Long
        cntVisible = Application.WorksheetFunction.CountA(hojaExtra.Range("A51:A58"))
        cntAutoCompletar = Application.WorksheetFunction.CountA(hojaExtra.Range("A60:A71"))
        
        ReDim listaVisible(1 To cntVisible)
        ReDim listaAutoCompletar(1 To cntVisible + cntAutoCompletar)
        
        i = 1
        j = 1
        For Each celda In hojaExtra.Range("A51:A58")
            If Trim(celda.Value) <> "" Then
                listaVisible(i) = celda.Value
                listaAutoCompletar(j) = celda.Value
                i = i + 1
                j = j + 1
            End If
        Next celda
        For Each celda In hojaExtra.Range("A60:A71")
            If Trim(celda.Value) <> "" Then
                listaAutoCompletar(j) = celda.Value
                j = j + 1
            End If
        Next celda
        
        For i = LBound(listaVisible) To UBound(listaVisible)
            cbxBuscarDato.AddItem listaVisible(i)
        Next i
        
        cbxBuscarDato.MatchEntry = fmMatchEntryComplete
    End If
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    If Sh.Name = "Buscar" Then
        With Sheets("Buscar")
            .Range("A3:P1048576").ClearContents
            .cbxFiltroDato.Value = ""
            .cbxBuscarDato.Value = ""
        End With
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    With Sheets("Buscar")
        .Range("A3:P1048576").ClearContents
        .cbxFiltroDato.Value = ""
        .cbxBuscarDato.Value = ""
    End With
End Sub

Private Sub cbxBuscarDato_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call btnBuscar_Click
    End If
End Sub

Private Sub cbxFiltroDato_Change()
    Dim wsFacturas As Worksheet
    Dim colIndex As Long
    Dim celda As Range
    Dim dict As Object
    Dim valores() As String
    Dim campo As String
    Dim i As Long, idx As Long
    Dim clave As Variant
    Dim valorFormateado As String

    Set dict = CreateObject("Scripting.Dictionary")
    Set wsFacturas = ThisWorkbook.Sheets("Facturas")

    cbxBuscarDato.Clear
    campo = cbxFiltroDato.Value
    If campo = "" Then Exit Sub

    ' Buscar índice de columna
    For i = 1 To 17
        If Trim(UCase(wsFacturas.Cells(1, i).Value)) = Trim(UCase(campo)) Then
            colIndex = i
            Exit For
        End If
    Next i

    If colIndex = 0 Then
        MsgBox "No se encontró el encabezado '" & campo & "' en la hoja Facturas.", vbExclamation
        Exit Sub
    End If

    ' Recopilar valores únicos en el diccionario
    For Each celda In wsFacturas.Range(wsFacturas.Cells(3, colIndex), wsFacturas.Cells(wsFacturas.Rows.count, colIndex).End(xlUp))
        If Trim(celda.Value) <> "" Then
              Select Case campo
                Case "FECHA DEL RECIBO"
                   If IsDate(celda.Value) Then
                      valorFormateado = Format(celda.Value, "dd/mm/yyyy hh:mm")
                      Else
                         valorFormateado = CStr(celda.Value)
                     End If
                  Case "FECHA DEL VUELO"
                    If IsDate(celda.Value) Then
                        valorFormateado = Format(celda.Value, "dd/mm/yyyy")
                    Else
                           valorFormateado = CStr(celda.Value)
                      End If
                  Case Else
                     valorFormateado = Trim(CStr(celda.Value))
             End Select
                     If Not dict.Exists(valorFormateado) Then
                            dict.Add valorFormateado, Nothing
                        End If
                 End If
              Next celda
    ' Pasar valores a array
    ReDim valores(0 To dict.count - 1)
    idx = 0
    For Each clave In dict.keys
        valores(idx) = clave
        idx = idx + 1
    Next clave

    ' Ordenar
    Call QuickSort(valores, LBound(valores), UBound(valores))

    ' Agregar al combo
    For i = LBound(valores) To UBound(valores)
        cbxBuscarDato.AddItem valores(i)
    Next i
End Sub

Private Sub btnLimpiar_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Buscar")
    
    Application.ScreenUpdating = False

    With ws
      .Range("A3:P1048576").ClearContents
    
      On Error Resume Next ' Evita error si no existen los controles
      .OLEObjects("cbxFiltroDato").Object.Value = ""
      .OLEObjects("cbxBuscarDato").Object.Value = ""
      On Error GoTo 0
    End With

    Application.ScreenUpdating = True
End Sub

Private Sub btnBuscar_Click()
    Dim wsFacturas As Worksheet, wsBuscar As Worksheet
    Dim campo As String, valorBuscado As String
    Dim iCol As Long, iFila As Long, filaDestino As Long
    Dim ultimaFilaFacturas As Long, ultimaFilaBuscar As Long
    Dim coincide As Boolean, valorCelda As Variant
    Dim esFecha As Boolean

    Set wsFacturas = ThisWorkbook.Sheets("Facturas")
    Set wsBuscar = ThisWorkbook.Sheets("Buscar")

    campo = Trim(cbxFiltroDato.Value)
    valorBuscado = Trim(cbxBuscarDato.Value)

    ' Si el campo está vacío, asumir "FECHA DE RECIBO"
    If campo = "" Then
      campo = "FECHA DE RECIBO"
    End If

    ' Si el valor está vacío, asumir "HOY"
    If valorBuscado = "" Then
        valorBuscado = "HOY"
    End If

    ' Obtener índice de columna en hoja "Facturas"
    iCol = 0
    For iCol = 1 To 17
        If Trim(UCase(wsFacturas.Cells(1, iCol).Value)) = Trim(UCase(campo)) Then Exit For
    Next iCol

    If iCol > 17 Then
        MsgBox "Campo no encontrado en hoja Facturas.", vbExclamation
        Exit Sub
    End If

    ' Determinar si el campo es de tipo fecha
    esFecha = (campo = "FECHA DE RECIBO" Or campo = "FECHA DEL VUELO")

    ' Limpiar datos anteriores
    ultimaFilaBuscar = wsBuscar.Cells(wsBuscar.Rows.count, "A").End(xlUp).row
    If ultimaFilaBuscar >= 3 Then
        wsBuscar.Range("A3:P" & ultimaFilaBuscar).ClearContents
    End If

    ' Procesar rangos de fechas especiales
    Dim hoy As Date: hoy = Date
    Dim valorClave As String: valorClave = UCase(valorBuscado)
    Dim fechaInicio As Date, fechaFin As Date

    Select Case valorClave
        Case "HOY"
            fechaInicio = hoy: fechaFin = hoy
        Case "AYER"
            fechaInicio = hoy - 1: fechaFin = hoy - 1
        Case "SEMANAL"
            fechaInicio = hoy - Weekday(hoy, vbMonday) + 1: fechaFin = hoy
        Case "MENSUAL"
            fechaInicio = DateSerial(Year(hoy), Month(hoy), 1)
            fechaFin = DateSerial(Year(hoy), Month(hoy) + 1, 0)
        Case "TRIMESTRE"
            fechaInicio = DateAdd("m", -3, hoy): fechaFin = hoy
        Case "SEMESTRE"
            fechaInicio = DateAdd("m", -6, hoy): fechaFin = hoy
        Case "ANUAL"
            fechaInicio = DateSerial(Year(hoy), 1, 1): fechaFin = hoy
        Case "TODO"
            fechaInicio = DateSerial(1900, 1, 1): fechaFin = DateSerial(2999, 12, 31)
        Case "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
            On Error Resume Next
            fechaInicio = DateSerial(Year(hoy), Month(DateValue("1/" & valorClave & "/" & Year(hoy))), 1)
            fechaFin = DateSerial(Year(hoy), Month(DateValue("1/" & valorClave & "/" & Year(hoy))) + 1, 0)
            On Error GoTo 0
        Case Else
            fechaInicio = 0: fechaFin = 0 ' Búsqueda por texto directo
    End Select

    ' Buscar y copiar coincidencias
    ultimaFilaFacturas = wsFacturas.Cells(wsFacturas.Rows.count, "A").End(xlUp).row
    filaDestino = 3

    For iFila = 3 To ultimaFilaFacturas
        coincide = False
        valorCelda = wsFacturas.Cells(iFila, iCol).Value

        If esFecha And fechaInicio <> 0 Then
            If IsDate(valorCelda) Then
                If DateValue(valorCelda) >= fechaInicio And DateValue(valorCelda) <= fechaFin Then
                    coincide = True
                End If
            End If
        ElseIf esFecha Then
            coincide = CoincideFecha(valorCelda, valorBuscado)
        Else
            If InStr(1, CStr(valorCelda), valorBuscado, vbTextCompare) > 0 Then
                coincide = True
            End If
        End If

        If coincide Then
            wsFacturas.Range(wsFacturas.Cells(iFila, 2), wsFacturas.Cells(iFila, 17)).Copy _
                Destination:=wsBuscar.Cells(filaDestino, 1)
            filaDestino = filaDestino + 1
        End If
    Next iFila

    If filaDestino = 3 Then
        MsgBox "No se encontraron coincidencias para el criterio especificado.", vbInformation
    End If
End Sub

Private Sub btnGenerar_Click()
    Dim ws As Worksheet
    Dim rutaCarpeta As String, nombreArchivo As String
    Dim filtro As String, valor As String
    Dim fechaActual As String
    Dim lastRow As Long, firstCol As Long, lastCol As Long
    Dim shp As Shape
    Dim col As Long
    Dim printRange As Range

    Set ws = ThisWorkbook.Sheets("Buscar")
    filtro = Replace(cbxFiltroDato.Value, ":", "-")
    valor = Replace(cbxBuscarDato.Value, ":", "-")
    valor = Replace(valor, "/", "-")

    fechaActual = Format(Now, "dd-mm-yyyy HH.mm")
    rutaCarpeta = ThisWorkbook.Path & "\PDFs Generados\"
    nombreArchivo = fechaActual & " - " & filtro & " - " & valor & " - Facturacion Administrativa CIAC.pdf"

    If Dir(rutaCarpeta, vbDirectory) = "" Then MkDir rutaCarpeta

    ' Configurar página
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperLetter
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = False
        .CenterVertically = False
        .RightFooter = "Generado el: " & Format(Now, "dd-mm-yyyy")
    End With

    ' Ocultar controles temporales
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Or shp.Type = msoOLEControlObject Then
            shp.visible = msoFalse
        End If
    Next shp

    ' Ocultar columnas O y P
    ws.Columns("O:P").Hidden = True

    ' Última fila con datos
    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row

    ' Detectar primera y última columna con datos
    firstCol = 0
    lastCol = 0
    For col = 1 To ws.Columns.count
        If Application.WorksheetFunction.CountA(ws.Columns(col)) > 0 Then
            If firstCol = 0 Then firstCol = col
            lastCol = col
        End If
    Next col

    If firstCol = 0 Or lastCol = 0 Then
        MsgBox "No hay datos para exportar."
        Exit Sub
    End If

    ' Establecer el área de impresión
    Set printRange = ws.Range(ws.Cells(1, firstCol), ws.Cells(lastRow, lastCol))
    ws.PageSetup.PrintArea = printRange.Address

    ' Exportar PDF
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=rutaCarpeta & nombreArchivo, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ' Mostrar columnas O y P nuevamente
    ws.Columns("O:P").Hidden = False

    ' Mostrar los controles otra vez
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Or shp.Type = msoOLEControlObject Then
            shp.visible = msoTrue
        End If
    Next shp

    MsgBox "PDF generado exitosamente en: " & rutaCarpeta & nombreArchivo
End Sub

Private Sub btnGuardar_Click()
    Dim wsFacturas As Worksheet, wsBuscar As Worksheet
    Dim filaBuscar As Long, filaFacturas As Long, ultimaFilaBuscar As Long
    Dim ID As Variant, celdaID As Range
    Dim fechaRecibo As Variant

    ' VALIDAR ANTES DE GUARDAR
    If Not ValidarDatosParaGuardar Then
        Exit Sub
    End If

    ' Proceder con la lógica de guardar
    Set wsFacturas = ThisWorkbook.Sheets("Facturas")
    Set wsBuscar = ThisWorkbook.Sheets("Buscar")
    
    ultimaFilaBuscar = wsBuscar.Cells(wsBuscar.Rows.count, "A").End(xlUp).row

    For filaBuscar = 3 To ultimaFilaBuscar
        fechaRecibo = wsBuscar.Cells(filaBuscar, 1).Value
        
        If IsDate(fechaRecibo) Then
            Set celdaID = Nothing
            For Each celdaID In wsFacturas.Range("B3:B" & wsFacturas.Cells(wsFacturas.Rows.count, "B").End(xlUp).row)
                If Format(celdaID.Value, "dd/mm/yyyy hh:mm") = Format(fechaRecibo, "dd/mm/yyyy hh:mm") Then
                    filaFacturas = celdaID.row
                    Exit For
                End If
            Next celdaID

            If filaFacturas > 0 Then
                wsBuscar.Range("A" & filaBuscar & ":P" & filaBuscar).Copy _
                    Destination:=wsFacturas.Range("B" & filaFacturas & ":Q" & filaFacturas)
            End If
        End If
    Next filaBuscar

    MsgBox "Los datos han sido guardados correctamente en la hoja Facturas.", vbInformation
End Sub

Private Function ValidarDatosParaGuardar() As Boolean
    Dim wsBuscar As Worksheet
    Dim fila As Long
    Dim valor As Variant
    Dim camposObligatorios As Variant
    Dim col As Variant
    Dim regexCedula As Object
    Dim celdaTexto As String

    Set wsBuscar = ThisWorkbook.Sheets("Buscar")
    Set regexCedula = CreateObject("VBScript.RegExp")
    regexCedula.IgnoreCase = True
    regexCedula.Global = False
    regexCedula.pattern = "^[VE]{1}\d{5,10}$" ' Ej: V33130224, E51546775

    camposObligatorios = Array(1, 2, 3, 5, 6, 9, 10) ' Índices obligatorios

    For fila = 3 To wsBuscar.Cells(wsBuscar.Rows.count, "A").End(xlUp).row
        For Each col In camposObligatorios
            valor = Trim(wsBuscar.Cells(fila, col).Value)
            If valor = "" Then
                MsgBox "Campo obligatorio vacío en la fila " & fila & ", columna '" & wsBuscar.Cells(2, col).Value & "'", vbExclamation
                ValidarDatosParaGuardar = False
                Exit Function
            End If
            ' Normalizar a mayúsculas si es texto
            If IsNumeric(valor) = False Then
                wsBuscar.Cells(fila, col).Value = UCase(valor)
            End If
        Next col

        ' Validación de fechas
        If Not IsDate(wsBuscar.Cells(fila, 1).Value) Then
            MsgBox "Formato inválido en 'FECHA DEL RECIBO' en la fila " & fila, vbExclamation
            ValidarDatosParaGuardar = False
            Exit Function
        End If
        If Not IsDate(wsBuscar.Cells(fila, 2).Value) Then
            MsgBox "Formato inválido en 'FECHA DEL VUELO' en la fila " & fila, vbExclamation
            ValidarDatosParaGuardar = False
            Exit Function
        End If

        ' Validación del combustible (si aplica)
        If wsBuscar.Cells(fila, 8).Value <> "" Then
            If Not IsNumeric(wsBuscar.Cells(fila, 8).Value) Or wsBuscar.Cells(fila, 8).Value < 0 Then
                MsgBox "Cantidad de combustible inválida en la fila " & fila, vbExclamation
                ValidarDatosParaGuardar = False
                Exit Function
            End If
        End If

        ' Validación de pago (columna 10)
        If Not IsNumeric(wsBuscar.Cells(fila, 10).Value) Or wsBuscar.Cells(fila, 10).Value < 0 Then
            MsgBox "Monto de PAGO inválido en la fila " & fila, vbExclamation
            ValidarDatosParaGuardar = False
            Exit Function
        End If

        ' Validación de cédula del alumno (col 6)
        celdaTexto = Trim(wsBuscar.Cells(fila, 6).Value)
        If Not regexCedula.Test(celdaTexto) Then
            MsgBox "Cédula del alumno inválida en la fila " & fila & ": " & celdaTexto, vbExclamation
            ValidarDatosParaGuardar = False
            Exit Function
        End If

        ' Validación de cédula del depositante (col 12) si no está vacía
        celdaTexto = Trim(wsBuscar.Cells(fila, 12).Value)
        If celdaTexto <> "" Then
            If Not regexCedula.Test(celdaTexto) Then
                MsgBox "Cédula del depositante inválida en la fila " & fila & ": " & celdaTexto, vbExclamation
                ValidarDatosParaGuardar = False
                Exit Function
            End If
        End If
    Next fila

    ValidarDatosParaGuardar = True
End Function

' Función para validar el formato de cédula
Private Function EsCedulaValida(cedula As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.pattern = "^[A-Z]\d{5,9}$" ' 1 letra + 5 a 9 dígitos
    regex.IgnoreCase = False
    regex.Global = False
    
    ' Eliminar espacios alrededor
    cedula = Trim(cedula)
    
    EsCedulaValida = regex.Test(cedula)
End Function

Private Function CoincideFecha(valorFecha As Variant, criterio As String) As Boolean
    On Error Resume Next
    CoincideFecha = False
    
    If Not IsDate(valorFecha) Then Exit Function
    
    Dim dt As Date: dt = CDate(valorFecha)
    criterio = Trim(LCase(criterio))
    Dim partes() As String
    
    ' Si contiene espacio, separar fecha y hora
    If InStr(criterio, " ") > 0 Then
        partes = Split(criterio, " ")
        If UBound(partes) = 1 Then
            ' Comparar fecha
            If IsDate(partes(0)) And Format(dt, "dd/mm/yyyy") = Format(CDate(partes(0)), "dd/mm/yyyy") Then
                ' Comparar hora
                If InStr(Format(dt, "HH:mm"), partes(1)) > 0 Then
                    CoincideFecha = True
                    Exit Function
                End If
            End If
        End If
    End If
    
    ' Solo hora
    If InStr(criterio, ":") > 0 Then
        If InStr(Format(dt, "HH:mm"), criterio) > 0 Then
            CoincideFecha = True
            Exit Function
        End If
    End If
    
    ' Fecha completa (dd/mm/yyyy)
    If IsDate(criterio) Then
        If Format(dt, "dd/mm/yyyy") = Format(CDate(criterio), "dd/mm/yyyy") Then
            CoincideFecha = True
            Exit Function
        End If
    End If

    ' mes/año
    If InStr(criterio, "/") > 0 Then
        partes = Split(criterio, "/")
        If UBound(partes) = 1 Then ' ej: 03/2025
            If IsNumeric(partes(0)) And IsNumeric(partes(1)) Then
                If Month(dt) = Val(partes(0)) And Year(dt) = Val(partes(1)) Then
                    CoincideFecha = True
                    Exit Function
                End If
            End If
        End If
    End If

    ' Solo año
    If Len(criterio) = 4 And IsNumeric(criterio) Then
        If Year(dt) = Val(criterio) Then
            CoincideFecha = True
            Exit Function
        End If
    End If

    ' Solo mes del año actual
    If IsNumeric(criterio) And Val(criterio) >= 1 And Val(criterio) <= 12 Then
        If Month(dt) = Val(criterio) And Year(dt) = Year(Date) Then
            CoincideFecha = True
            Exit Function
        End If
    End If
End Function

-------------------------------------------------------------------------------