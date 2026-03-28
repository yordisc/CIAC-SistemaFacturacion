Option Explicit

Private Sub Worksheet_Activate()
    Call InicializarControles
End Sub

Public Sub InicializarControles()
    On Error GoTo ErrorHandler

    Dim wsBuscar As Worksheet
    Dim primerFuncionario As String
    Dim i As Long
    Dim wsDatos As Worksheet

    ' Referenciar la hoja "Buscar"
    Set wsBuscar = ThisWorkbook.Sheets("Buscar")
    Set wsDatos = ThisWorkbook.Sheets("Datos")
    
    ' Cargar datos en ComboBoxes
    CargarComboUsuarios "FUNCIONARIO", cbxDespachador
    CargarComboUsuarios "INSTRUCTOR", cbxInstructor
    CargarFactura cbxFactura
    CargarAeronaves cbxAeronave
    CargarMontos cbxMonto
    CargarMetodosPago cbxMPago

    ' Configurar propiedades
    cbxMPago.Enabled = False

    ' Limpiar campos
    LimpiarCampos
    Sheets("Factura").OLEObjects("cbxFactura").Object.Text = "SELECCIONA"

    ' Asignar fecha actual por defecto
    Dim fechaActual As String
    fechaActual = Format(Date, "dd/mm/yyyy")
    
    If Not Sheets("Factura").OLEObjects("cbxFecha") Is Nothing Then
        CargarFechaActual Sheets("Factura").OLEObjects("cbxFecha").Object
    Else
       MsgBox "Error: el ComboBox 'cbxFecha' no se encuentra disponible.", vbExclamation
    End If

    ' Si la hoja "Buscar" está oculta, hacer configuraciones especiales
    If wsBuscar.visible = xlSheetHidden Or wsBuscar.visible = xlSheetVeryHidden Then
        
        ' 1. Bloquear cbxFactura en "COMBUSTIBLE"
        With Sheets("Factura").OLEObjects("cbxFactura").Object
            .Clear
            .AddItem "COMBUSTIBLE"
            .ListIndex = 0
            .Enabled = False
        End With
        
        ' 2. Excluir primer funcionario de cbxDespachador
        ' Obtener primer funcionario
        primerFuncionario = Trim(wsDatos.Cells(2, 3).Value & " " & wsDatos.Cells(2, 4).Value)
        
        ' Volver a cargar el cbxDespachador SIN el primer funcionario
        cbxDespachador.Clear
        For i = 2 To wsDatos.Cells(wsDatos.Rows.count, "A").End(xlUp).row
            If wsDatos.Cells(i, 2).Value = "FUNCIONARIO" Then
                Dim nombreCompleto As String
                nombreCompleto = Trim(wsDatos.Cells(i, 3).Value & " " & wsDatos.Cells(i, 4).Value)
                If EsValidoNombreApellido(nombreCompleto) Then
                    If nombreCompleto <> primerFuncionario Then
                        cbxDespachador.AddItem nombreCompleto
                    End If
                End If
            End If
        Next i
    Else
        ' Si "Buscar" está visible, asegurar que cbxFactura esté habilitado
        Sheets("Factura").OLEObjects("cbxFactura").Object.Enabled = True
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Ha ocurrido un error: " & Err.Description
End Sub

Private Sub CargarEnComboBox(hoja As String, rango As String, cbx As ComboBox)
    Dim ws As Worksheet
    Dim celda As Range

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Sheets(hoja)
    cbx.Clear

    For Each celda In ws.Range(rango)
        If Trim(celda.Value) <> "" Then
            cbx.AddItem Trim(celda.Value)
        End If
    Next celda

    Exit Sub

ErrorHandler:
    MsgBox "Error al cargar el ComboBox: " & Err.Description, vbCritical
End Sub
Private Sub CargarFechaActual(cbx As ComboBox)
    cbx.Clear
    cbx.AddItem Format(Date, "dd/mm/yyyy")
End Sub
Private Sub CargarComboUsuarios(cargo As String, cbx As ComboBox)
    Dim ws As Worksheet
    Dim i As Long
    Dim nombreCompleto As String
    
    Set ws = ThisWorkbook.Worksheets("Datos")
    cbx.Clear
    
    For i = 2 To ws.Cells(ws.Rows.count, "A").End(xlUp).row
        If ws.Cells(i, 2).Value = cargo Then
            nombreCompleto = Trim(ws.Cells(i, 3).Value & " " & ws.Cells(i, 4).Value)
            If EsValidoNombreApellido(nombreCompleto) Then
                cbx.AddItem nombreCompleto
            End If
        End If
    Next i
End Sub
Private Sub CargarCedulasAlumnos(cbx As ComboBox)
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("Datos")
    cbx.Clear
    
    For i = 2 To ws.Cells(ws.Rows.count, "A").End(xlUp).row
        If ws.Cells(i, 2).Value = "ALUMNO" Then
            If EsCedulaValida(ws.Cells(i, 5).Value) Then
                cbx.AddItem ws.Cells(i, 5).Value
            End If
        End If
    Next i
End Sub
Private Sub CargarAeronaves(cbx As ComboBox)
    Dim ws As Worksheet
    Dim i As Long
    Dim dict As Object
    
    Set ws = ThisWorkbook.Worksheets("Facturas")
    Set dict = CreateObject("Scripting.Dictionary")
    cbx.Clear
    
    For i = 2 To ws.Cells(ws.Rows.count, "A").End(xlUp).row
        If Not dict.Exists(ws.Cells(i, 8).Value) And ws.Cells(i, 8).Value <> "" Then
            dict.Add ws.Cells(i, 8).Value, 1
            cbx.AddItem ws.Cells(i, 8).Value
        End If
    Next i
End Sub
Private Sub CargarFactura(cbx As ComboBox)
    CargarEnComboBox "Extras", "A45:A48", cbx
End Sub

Private Sub CargarMontos(cbx As ComboBox)
    CargarEnComboBox "Extras", "A7:A8", cbx
End Sub
Private Sub CargarMetodosPago(cbx As ComboBox)
    CargarEnComboBox "Extras", "A11:A13", cbx
End Sub
Private Sub CargarBancos(cbx As ComboBox)
    CargarEnComboBox "Extras", "B17:B43", cbx
End Sub
Private Sub CargarCodigosBancos(cbx As ComboBox)
    CargarEnComboBox "Extras", "A17:A43", cbx
End Sub

' ================= EVENTOS DE CONTROLES =================

' Botón Guardar
Private Sub btnGuardar_Click()
    If ValidarDatos() Then
        GuardarFactura
        LimpiarCampos
        MsgBox "Factura guardada correctamente.", vbInformation
    End If
End Sub
' Botón Limpiar
Private Sub btnLimpiar_Click()
    LimpiarCampos
End Sub

Private Sub cbxBanco_Change()
    SincronizarCodigoPorBanco cbxBanco.Text
End Sub
Private Sub cbxBCodigo_Change()
    SincronizarBancoPorCodigo cbxBCodigo.Text
End Sub
Private Sub cbxFactura_Change()
    Dim hoja As Worksheet
    Set hoja = Sheets("Factura")

    ' Lista de todos los controles
    Dim todosControles As Variant
    todosControles = Array("cbxFecha", "cbxDespachador", "cbxInstructor", "cbxAlumno", "cbxCAlumno", _
                           "cbxObservacion", "cbxMonto", "cbxMPago", "cbxAeronave", "cbxLitros", _
                           "cbxBanco", "cbxBCodigo", "cbxCedulaD", "cbxNumOperacion", "cbxNumTlOrigen", _
                           "cbxCantidad", "btnLimpiar", "btnGuardar")

    Dim mostrarControles As Collection
    Set mostrarControles = New Collection

    ' Ocultar todos los textos por defecto
    hoja.Range("B2:F19").Font.color = RGB(142, 169, 219)
    hoja.Range("D11").Font.color = RGB(142, 169, 219) ' Asegurar que D11 inicie oculta
    hoja.OLEObjects("cbxCantidad").visible = False ' Asegurar que el control también esté oculto inicialmente

    ' Mostrar siempre cbxFactura
    hoja.OLEObjects("cbxFactura").visible = True

    Select Case UCase(cbxFactura.Value)
        Case "SELECCIONA"
            ' Nada más se oculta todo

        Case "HONORARIO"
            mostrarControles.Add "cbxFecha"
            mostrarControles.Add "cbxDespachador"
            mostrarControles.Add "cbxInstructor"
            mostrarControles.Add "cbxAlumno"
            mostrarControles.Add "cbxCAlumno"
            mostrarControles.Add "cbxObservacion"
            mostrarControles.Add "cbxMonto"
            mostrarControles.Add "cbxMPago"
            mostrarControles.Add "btnLimpiar"
            mostrarControles.Add "btnGuardar"

            hoja.Range("B2:F7").Font.color = RGB(0, 0, 0)
            hoja.Range("B10:F11").Font.color = RGB(0, 0, 0)

            hoja.OLEObjects("cbxAeronave").Object.Text = "NO APLICA"
            hoja.OLEObjects("cbxLitros").Object.Text = "NO APLICA"
            hoja.OLEObjects("cbxMonto").Object.Text = ""
            hoja.OLEObjects("cbxMPago").Object.Text = ""

        Case "COMBUSTIBLE"
            mostrarControles.Add "cbxFecha"
            mostrarControles.Add "cbxDespachador"
            mostrarControles.Add "cbxAlumno"
            mostrarControles.Add "cbxCAlumno"
            mostrarControles.Add "cbxAeronave"
            mostrarControles.Add "cbxLitros"
            mostrarControles.Add "cbxObservacion"
            mostrarControles.Add "cbxMonto"
            mostrarControles.Add "cbxMPago"
            mostrarControles.Add "btnLimpiar"
            mostrarControles.Add "btnGuardar"

            hoja.Range("B2:F11").Font.color = RGB(0, 0, 0)
            hoja.Range("B5").Font.color = RGB(142, 169, 219)

            hoja.OLEObjects("cbxInstructor").Object.Text = "NO APLICA"
            hoja.OLEObjects("cbxMonto").Object.Text = ""
            hoja.OLEObjects("cbxMPago").Object.Text = ""

        Case "H&C"
            mostrarControles.Add "cbxFecha"
            mostrarControles.Add "cbxDespachador"
            mostrarControles.Add "cbxAlumno"
            mostrarControles.Add "cbxCAlumno"
            mostrarControles.Add "cbxInstructor"
            mostrarControles.Add "cbxAeronave"
            mostrarControles.Add "cbxLitros"
            mostrarControles.Add "cbxObservacion"
            mostrarControles.Add "cbxMonto"
            mostrarControles.Add "cbxMPago"
            mostrarControles.Add "btnLimpiar"
            mostrarControles.Add "btnGuardar"

            hoja.Range("B2:F11").Font.color = RGB(0, 0, 0)
            hoja.OLEObjects("cbxMonto").Object.Text = ""
            hoja.OLEObjects("cbxMPago").Object.Text = ""
    End Select

    ' Ocultar todos los controles
    Dim nombre As Variant
    For Each nombre In todosControles
        hoja.OLEObjects(nombre).visible = False
    Next nombre

    ' Mostrar solo los controles relevantes
    For Each nombre In mostrarControles
        hoja.OLEObjects(nombre).visible = True
    Next nombre

    ' Aplicar lógica de visibilidad para D11 según función
    Call ActualizarVisibilidadCantidad
End Sub
' Cambio en cbxMonto
Private Sub cbxMonto_Change()

    Dim hoja As Worksheet
    Set hoja = Sheets("Factura")

    If UCase(cbxMonto.Text) = "DIVISAS" Then
        cbxMPago.Text = "EFECTIVO"
        cbxMPago.Enabled = False
        cbxCantidad.Enabled = True
        MostrarCamposPagoMovil False
        hoja.Range("B13:F17").Font.color = RGB(142, 169, 219)
        cbxBanco.Clear
        cbxBCodigo.Clear
    ElseIf UCase(cbxMonto.Text) = "BOLIVARES" Then
        cbxMPago.Text = "PAGOMOVIL"
        cbxMPago.Enabled = True
        cbxCantidad.Enabled = True
    End If
        Call ActualizarVisibilidadCantidad
End Sub
' Cambio en cbxMPago
Private Sub cbxMPago_Change()

    Dim hoja As Worksheet
    Set hoja = Sheets("Factura")

    If cbxMPago.Text = "PAGOMOVIL" And UCase(cbxMonto.Text) = "BOLIVARES" Then
        CargarBancos cbxBanco
        CargarCodigosBancos cbxBCodigo
        MostrarCamposPagoMovil True
        hoja.Range("B13:F17").Font.color = RGB(0, 0, 0)
    Else
        cbxBanco.Clear
        cbxBCodigo.Clear
        cbxCantidad.Clear
        cbxCantidad.Enabled = False
        MostrarCamposPagoMovil False
        hoja.Range("B13:F17").Font.color = RGB(142, 169, 219)
    End If
        Call ActualizarVisibilidadCantidad
End Sub

Private Sub cbxBanco_Enter()
    If Not (cbxMonto.Text = "BOLIVARES" And cbxMPago.Text = "PAGOMOVIL") Then
        cbxBanco.ListRows = 0
        cbxBanco.Clear
    End If
End Sub
Private Sub cbxBCodigo_Enter()
    If Not (cbxMonto.Text = "BOLIVARES" And cbxMPago.Text = "PAGOMOVIL") Then
        cbxBCodigo.ListRows = 0
        cbxBCodigo.Clear
    End If
End Sub

' Validar al salir de cbxDespachador
Private Sub cbxDespachador_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not EsValidoNombreApellido(cbxDespachador.Text) Then
        MsgBox "Debe ingresar nombre y apellido del despachador.", vbExclamation
        Cancel = True
    End If
End Sub
' Validar al salir de cbxInstructor
Private Sub cbxInstructor_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not EsValidoNombreApellido(cbxInstructor.Text) Then
        MsgBox "Debe ingresar nombre y apellido del instructor.", vbExclamation
        Cancel = True
    End If
End Sub
' Validar al salir de cbxAlumno
Private Sub cbxAlumno_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not EsValidoNombreApellido(cbxAlumno.Text) Then
        MsgBox "Debe ingresar nombre y apellido del alumno.", vbExclamation
        Cancel = True
    End If
End Sub
' Validar al salir de cbxCAlumno
Private Sub cbxCAlumno_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not EsCedulaValida(cbxCAlumno.Text) Then
        MsgBox "Cédula inválida. Formato: V12345678", vbExclamation
        Cancel = True
    Else
        cbxCAlumno.Text = FormatearCedula(cbxCAlumno.Text)
    End If
End Sub

' Eventos Click para todos los ComboBoxes
Private Sub cbxAlumno_Click()
    If cbxAlumno.ListCount > 0 Then cbxAlumno.DropDown
End Sub
Private Sub cbxAeronave_Click()
    If cbxAeronave.ListCount > 0 Then cbxAeronave.DropDown
End Sub
Private Sub cbxCAlumno_Click()
    If cbxCAlumno.ListCount > 0 Then cbxCAlumno.DropDown
End Sub
Private Sub cbxCedulaD_Click()
    If cbxCedulaD.ListCount > 0 Then cbxCedulaD.DropDown
End Sub
Private Sub cbxDespachador_Click()
    If cbxDespachador.ListCount > 0 Then cbxDespachador.DropDown
End Sub
Private Sub cbxInstructor_Click()
    If cbxInstructor.ListCount > 0 Then cbxInstructor.DropDown
End Sub
Private Sub cbxMonto_Click()
    If cbxMonto.ListCount > 0 Then cbxMonto.DropDown
End Sub
Private Sub cbxMPago_Click()
    If cbxMPago.ListCount > 0 Then cbxMPago.DropDown
End Sub
Private Sub cbxBanco_Click()
    If cbxBanco.ListCount > 0 Then cbxBanco.DropDown
End Sub
Private Sub cbxBCodigo_Click()
    ' Prevenir que se abra el menú contextual
    cbxBCodigo.ListRows = 0
    ' Cancelar cualquier apertura de lista
End Sub

Private Sub cbxBCodigo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Redundancia extra para prevenir el despliegue
    cbxBCodigo.ListRows = 0
End Sub

' ================= FUNCIONES PRIVADAS =================

Private Sub ActualizarVisibilidadCantidad()
    Dim hoja As Worksheet
    Set hoja = Sheets("Factura")

    Dim montoValido As Boolean
    Dim pagoValido As Boolean

    montoValido = (UCase(cbxMonto.Text) = "BOLIVARES" Or UCase(cbxMonto.Text) = "DIVISAS")
    pagoValido = (UCase(cbxMPago.Text) = "PAGOMOVIL" Or UCase(cbxMPago.Text) = "EFECTIVO")

    If montoValido And pagoValido Then
        hoja.Range("D11").Font.color = RGB(0, 0, 0)
        cbxCantidad.Enabled = True
        hoja.OLEObjects("cbxCantidad").visible = True
    Else
        hoja.Range("D11").Font.color = RGB(142, 169, 219)
        cbxCantidad.Clear
        cbxCantidad.Enabled = False
        hoja.OLEObjects("cbxCantidad").visible = False
    End If
End Sub

' Validar todos los datos antes de guardar
Private Function ValidarDatos() As Boolean
    On Error GoTo ErrorHandler
    Dim hoja As Worksheet
    Set hoja = Sheets("Factura")
    ValidarDatos = False

    ' Atajos para acceder al texto de cada control
    Dim despachador As String: despachador = UCase(Trim(hoja.OLEObjects("cbxDespachador").Object.Text))
    Dim instructor As String: instructor = UCase(Trim(hoja.OLEObjects("cbxInstructor").Object.Text))
    Dim alumno As String: alumno = UCase(Trim(hoja.OLEObjects("cbxAlumno").Object.Text))
    Dim cedulaAlumno As String: cedulaAlumno = UCase(Trim(hoja.OLEObjects("cbxCAlumno").Object.Text))
    Dim aeronave As String: aeronave = UCase(Trim(hoja.OLEObjects("cbxAeronave").Object.Text))
    Dim litros As String: litros = UCase(Trim(hoja.OLEObjects("cbxLitros").Object.Text))
    Dim monto As String: monto = UCase(Trim(hoja.OLEObjects("cbxMonto").Object.Text))
    Dim metodoPago As String: metodoPago = UCase(Trim(hoja.OLEObjects("cbxMPago").Object.Text))
    Dim tipoFactura As String: tipoFactura = UCase(Trim(hoja.OLEObjects("cbxFactura").Object.Text))

    ' Campos de Pago Móvil
    Dim banco As String: banco = UCase(Trim(hoja.OLEObjects("cbxBanco").Object.Text))
    Dim codigoBanco As String: codigoBanco = UCase(Trim(hoja.OLEObjects("cbxBCodigo").Object.Text))
    Dim cedulaDep As String: cedulaDep = UCase(Trim(hoja.OLEObjects("cbxCedulaD").Object.Text))
    Dim numOperacion As String: numOperacion = UCase(Trim(hoja.OLEObjects("cbxNumOperacion").Object.Text))
    Dim numTlOrigen As String: numTlOrigen = UCase(Trim(hoja.OLEObjects("cbxNumTlOrigen").Object.Text))

    ' === Validar tipo de factura ===
    If tipoFactura <> "HONORARIO" And tipoFactura <> "COMBUSTIBLE" And tipoFactura <> "H&C" Then
        MsgBox "Tipo de factura inválido. Solo se permiten: HONORARIOS, COMBUSTIBLE o H&C.", vbCritical
        hoja.OLEObjects("cbxFactura").Activate
        Exit Function
    End If

    ' === Exclusiones según el tipo de factura ===
    If tipoFactura = "HONORARIO" Then
        aeronave = ""
        litros = ""
    End If

    If tipoFactura = "COMBUSTIBLE" Then
        instructor = ""
    End If

    ' === Exclusiones si método de pago es EFECTIVO ===
    If metodoPago = "EFECTIVO" Then
        banco = ""
        codigoBanco = ""
        cedulaDep = ""
        numOperacion = ""
        numTlOrigen = ""
    End If

    ' Validar campos obligatorios
    If despachador = "" Or alumno = "" Or cedulaAlumno = "" Or monto = "" Or metodoPago = "" Then
        MsgBox "Debe completar todos los campos obligatorios. (Debes apegarte a este formato: V123456789)", vbExclamation
        Exit Function
    End If

    ' Validaciones específicas
    If Not EsValidoNombreApellido(despachador) Then
        MsgBox "Nombre de despachador inválido.(Solo Funcionarios registrados)", vbExclamation
        hoja.OLEObjects("cbxDespachador").Activate
        Exit Function
    End If

    If tipoFactura <> "COMBUSTIBLE" Then
        If instructor = "" Or Not EsValidoNombreApellido(instructor) Then
            MsgBox "Nombre de instructor inválido. (Solo primer nombre y primer apellido)", vbExclamation
            hoja.OLEObjects("cbxInstructor").Activate
            Exit Function
        End If
    End If

    If Not EsValidoNombreApellido(alumno) Then
        MsgBox "Nombre de alumno inválido. (Solo primer nombre y primer apellido)", vbExclamation
        hoja.OLEObjects("cbxAlumno").Activate
        Exit Function
    End If

    If Not EsCedulaValida(cedulaAlumno) Then
        MsgBox "Cédula de alumno inválida. (Debes apegarte a este formato: V123456789)", vbExclamation
        hoja.OLEObjects("cbxCAlumno").Activate
        Exit Function
    End If

    If tipoFactura <> "HONORARIO" Then
        If aeronave = "" Then
            MsgBox "Debe seleccionar una aeronave.", vbExclamation
            hoja.OLEObjects("cbxAeronave").Activate
            Exit Function
        End If
        If Not IsNumeric(litros) Then
            MsgBox "Cantidad de combustible debe ser numérica.", vbExclamation
            hoja.OLEObjects("cbxLitros").Activate
            Exit Function
        End If
    End If

    If Not EsMontoValido(monto) Then
        MsgBox "Tipo de monto inválido.", vbExclamation
        hoja.OLEObjects("cbxMonto").Activate
        Exit Function
    End If

    If Not EsMetodoPagoValido(metodoPago) Then
        MsgBox "Método de pago inválido.", vbExclamation
        hoja.OLEObjects("cbxMPago").Activate
        Exit Function
    End If

    ' Validar si el método de pago es Pago Móvil
    If metodoPago = "PAGOMOVIL" Then
        If banco = "" Then
            MsgBox "Debe seleccionar un banco.", vbExclamation
            hoja.OLEObjects("cbxBanco").Activate
            Exit Function
        End If

        If codigoBanco = "" Then
            MsgBox "Debe seleccionar un código de banco.", vbExclamation
            hoja.OLEObjects("cbxBCodigo").Activate
            Exit Function
        End If

        If Not IsNumeric(numTlOrigen) Or Len(numTlOrigen) <> 11 Then
            MsgBox "Número telefónico debe ser numérico y tener exactamente 11 dígitos.", vbExclamation
            hoja.OLEObjects("cbxNumTlOrigen").Activate
            Exit Function
        End If

        If Not EsCedulaValida(cedulaDep) Then
            MsgBox "Cédula de depositante inválida.(Debes apegarte a este formato: V123456789)", vbExclamation
            hoja.OLEObjects("cbxCedulaD").Activate
            Exit Function
        End If

        If Not IsNumeric(numOperacion) Then
            MsgBox "Número de operación debe ser numérico.", vbExclamation
            hoja.OLEObjects("cbxNumOperacion").Activate
            Exit Function
        End If
    End If

    ValidarDatos = True
    Exit Function

ErrorHandler:
    MsgBox "Error al validar datos: " & Err.Description, vbCritical
End Function

' Guardar factura en la hoja Facturas
Private Sub GuardarFactura()
    Dim wsFacturas As Worksheet
    Set wsFacturas = ThisWorkbook.Worksheets("Facturas")

    Dim wsDatos As Worksheet
    Set wsDatos = ThisWorkbook.Worksheets("Datos")

    Dim tipoFactura As String
    tipoFactura = UCase(Trim(cbxFactura.Text))

    If tipoFactura <> "HONORARIO" And tipoFactura <> "COMBUSTIBLE" And tipoFactura <> "H&C" Then
        MsgBox "Tipo de factura inválido. Solo se permiten: HONORARIOS, COMBUSTIBLE o H&C.", vbCritical
        Exit Sub
    End If

    ' Buscar siguiente fila vacía en Facturas
    Dim fila As Long
    fila = wsFacturas.Cells(wsFacturas.Rows.count, "A").End(xlUp).row + 1

    ' Obtener y formatear la fecha
    Dim fechaFactura As String
    If Trim(cbxFecha.Text) = "" Then
        fechaFactura = Format(Date, "dd/mm/yyyy")
    Else
        fechaFactura = UCase(Trim(cbxFecha.Text))
    End If

    ' ------------------------
    ' Procesar nombre INSTRUCTOR
    ' ------------------------
    Dim textoInstructor As String
    Dim nombreInstructor As String, apellidoInstructor As String
    textoInstructor = UCase(Trim(cbxInstructor.Text))

    If textoInstructor <> "" Then
        Dim partesI() As String
        partesI = Split(textoInstructor, " ")
        If UBound(partesI) >= 1 Then
            nombreInstructor = partesI(0)
            apellidoInstructor = ""
            Dim j As Integer
            For j = 1 To UBound(partesI)
                apellidoInstructor = apellidoInstructor & partesI(j) & " "
            Next j
            apellidoInstructor = Trim(apellidoInstructor)
        Else
            nombreInstructor = textoInstructor
            apellidoInstructor = ""
        End If
    Else
        nombreInstructor = ""
        apellidoInstructor = ""
    End If

    ' ------------------------
    ' Procesar nombre ALUMNO
    ' ------------------------
    Dim textoAlumno As String
    Dim nombreAlumno As String, apellidoAlumno As String
    textoAlumno = UCase(Trim(cbxAlumno.Text))

    If textoAlumno <> "" Then
        Dim partesA() As String
        partesA = Split(textoAlumno, " ")
        If UBound(partesA) >= 1 Then
            nombreAlumno = partesA(0)
            apellidoAlumno = ""
            Dim k As Integer
            For k = 1 To UBound(partesA)
                apellidoAlumno = apellidoAlumno & partesA(k) & " "
            Next k
            apellidoAlumno = Trim(apellidoAlumno)
        Else
            nombreAlumno = textoAlumno
            apellidoAlumno = ""
        End If
    Else
        nombreAlumno = ""
        apellidoAlumno = ""
    End If

    Dim cedulaAlumno As String: cedulaAlumno = UCase(Trim(cbxCAlumno.Text))
    Dim cargoI As String: cargoI = "INSTRUCTOR"
    Dim cargoA As String: cargoA = "ALUMNO"

    ' Guardar datos en Facturas
    wsFacturas.Cells(fila, 1).Value = GenerarIDUnico()
    wsFacturas.Cells(fila, 2).Value = Now
    wsFacturas.Cells(fila, 3).Value = fechaFactura
    wsFacturas.Cells(fila, 4).Value = UCase(Trim(cbxDespachador.Text))

    If tipoFactura <> "COMBUSTIBLE" Then
        wsFacturas.Cells(fila, 5).Value = nombreInstructor & " " & apellidoInstructor
    Else
        wsFacturas.Cells(fila, 5).Value = ""
    End If

    wsFacturas.Cells(fila, 6).Value = nombreAlumno & " " & apellidoAlumno
    wsFacturas.Cells(fila, 7).Value = cedulaAlumno

    If tipoFactura <> "HONORARIO" Then
        wsFacturas.Cells(fila, 8).Value = UCase(Trim(cbxAeronave.Text))
        wsFacturas.Cells(fila, 9).Value = Val(cbxLitros.Text)
    Else
        wsFacturas.Cells(fila, 8).Value = ""
        wsFacturas.Cells(fila, 9).Value = ""
    End If

    wsFacturas.Cells(fila, 10).Value = UCase(Trim(cbxMonto.Text))
    wsFacturas.Cells(fila, 11).Value = UCase(Trim(cbxCantidad.Text))

    If UCase(Trim(cbxMPago.Text)) = "PAGOMOVIL" Then
        wsFacturas.Cells(fila, 12).Value = UCase(Trim(cbxBCodigo.Text))
        wsFacturas.Cells(fila, 13).Value = UCase(Trim(cbxCedulaD.Text))
        wsFacturas.Cells(fila, 14).Value = UCase(Trim(cbxNumOperacion.Text))
        wsFacturas.Cells(fila, 15).Value = UCase(Trim(cbxNumTlOrigen.Text))
    Else
        wsFacturas.Cells(fila, 12).Value = ""
        wsFacturas.Cells(fila, 13).Value = ""
        wsFacturas.Cells(fila, 14).Value = ""
        wsFacturas.Cells(fila, 15).Value = ""
    End If

    wsFacturas.Cells(fila, 16).Value = tipoFactura
    wsFacturas.Cells(fila, 17).Value = UCase(Trim(cbxObservacion.Text))

    ' --- Guardar en hoja Datos ---

    ' Instructor
    If tipoFactura <> "COMBUSTIBLE" And nombreInstructor <> "" Then
        If Not ExisteEnDatos(wsDatos, nombreInstructor, apellidoInstructor, "NO APLICA", cargoI) Then
            Dim filaDatosI As Long
            filaDatosI = wsDatos.Cells(wsDatos.Rows.count, "A").End(xlUp).row + 1
            wsDatos.Cells(filaDatosI, 1).Value = GenerarIDUnico()
            wsDatos.Cells(filaDatosI, 2).Value = cargoI
            wsDatos.Cells(filaDatosI, 3).Value = nombreInstructor
            wsDatos.Cells(filaDatosI, 4).Value = apellidoInstructor
            wsDatos.Cells(filaDatosI, 5).Value = "NO APLICA"
        End If
    End If

    ' Alumno
    If nombreAlumno <> "" And cedulaAlumno <> "" Then
        If Not ExisteEnDatos(wsDatos, nombreAlumno, apellidoAlumno, cedulaAlumno, cargoA) Then
            Dim filaDatosA As Long
            filaDatosA = wsDatos.Cells(wsDatos.Rows.count, "A").End(xlUp).row + 1
            wsDatos.Cells(filaDatosA, 1).Value = GenerarIDUnico()
            wsDatos.Cells(filaDatosA, 2).Value = cargoA
            wsDatos.Cells(filaDatosA, 3).Value = nombreAlumno
            wsDatos.Cells(filaDatosA, 4).Value = apellidoAlumno
            wsDatos.Cells(filaDatosA, 5).Value = cedulaAlumno
        End If
    End If
End Sub

'función para verificar duplicados
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

Private Function BuscarMejorCoincidencia(valorBuscado As String, cbx As ComboBox) As String
    Dim i As Long
    Dim item As String
    Dim mejorCoincidencia As String
    Dim mejorPuntaje As Long
    Dim puntajeActual As Long

    valorBuscado = LCase(Trim(valorBuscado))

    For i = 0 To cbx.ListCount - 1
        item = LCase(cbx.List(i))
        If item Like "*" & valorBuscado & "*" Then
            puntajeActual = Len(valorBuscado)
            If puntajeActual > mejorPuntaje Then
                mejorPuntaje = puntajeActual
                mejorCoincidencia = cbx.List(i)
            End If
        End If
    Next i

    BuscarMejorCoincidencia = mejorCoincidencia
End Function

Private Sub mostrarControles(nombres() As String)
    Dim nombre As Variant
    For Each nombre In nombres
        ThisWorkbook.Sheets("Factura").OLEObjects(nombre).visible = True
    Next nombre
End Sub

Private Sub CargarFechas(cbx As ComboBox)
    Dim i As Integer
    cbx.Clear
    For i = 0 To 30 ' Cargar los últimos 30 días
        cbx.AddItem Format(Date - i, "dd/mm/yyyy")
    Next i
End Sub

' Sincronizar nombre del banco al escribir código
Private Sub SincronizarBancoPorCodigo(codigo As String)
    If Trim(codigo) = "" Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Extras")

    Dim i As Long
    For i = 17 To 43
        If Trim(ws.Cells(i, 1).Value) = Trim(codigo) Then
            cbxBanco.Text = Trim(ws.Cells(i, 2).Value)
            Exit For
        End If
    Next i
End Sub

' Sincronizar código al seleccionar banco
Private Sub SincronizarCodigoPorBanco(nombreBanco As String)
    If Trim(nombreBanco) = "" Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Extras")

    Dim i As Long
    For i = 17 To 43
        If Trim(ws.Cells(i, 2).Value) = Trim(nombreBanco) Then
            cbxBCodigo.Text = Trim(ws.Cells(i, 1).Value)
            Exit For
        End If
    Next i
End Sub

Private Sub SincronizarAlumnoCedula()
    Dim i As Long
    Dim nombre As String
    Dim cedula As String
    Dim encontrado As Boolean
    
    nombre = Trim(cbxAlumno.Value)
    encontrado = False

    ' Buscar el nombre en la lista
    For i = 0 To cbxAlumno.ListCount - 1
        If StrComp(cbxAlumno.List(i), nombre, vbTextCompare) = 0 Then
            cbxCAlumno.Value = cbxCAlumno.List(i)
            encontrado = True
            Exit For
        End If
    Next i
    
    ' Si no lo encuentra, limpia la cédula
    If Not encontrado Then
        cbxCAlumno.Value = ""
    End If
End Sub

Private Sub SincronizarCedulaAlumno()
    Dim i As Long
    Dim cedula As String
    Dim nombre As String
    Dim encontrado As Boolean
    
    cedula = UCase(Trim(cbxCAlumno.Value)) ' Asegura formato V12345678 o E12345678
    encontrado = False

    ' Buscar la cédula en la lista
    For i = 0 To cbxCAlumno.ListCount - 1
        If StrComp(cbxCAlumno.List(i), cedula, vbTextCompare) = 0 Then
            cbxAlumno.Value = cbxAlumno.List(i)
            encontrado = True
            Exit For
        End If
    Next i
    
    ' Si no lo encuentra, limpia el nombre
    If Not encontrado Then
        cbxAlumno.Value = ""
    End If
End Sub

Private Sub MoverAlSiguienteControl(controlActual As Object)
    Dim controlesVisibles As Collection
    Dim ctrl As OLEObject
    Dim i As Long

    Set controlesVisibles = New Collection

    ' Recolectar solo controles relevantes y visibles/habilitados
    For Each ctrl In Me.OLEObjects
        On Error Resume Next
        If (TypeName(ctrl.Object) = "ComboBox" Or TypeName(ctrl.Object) = "CommandButton") Then
            If ctrl.Object.visible And ctrl.Object.Enabled Then
                controlesVisibles.Add ctrl.Object
            End If
        End If
        On Error GoTo 0
    Next ctrl

    ' Buscar el índice del control actual comparando por objeto
    For i = 1 To controlesVisibles.count
        If controlesVisibles(i) Is controlActual Then Exit For
    Next i

    ' Pasar al siguiente control, si existe
    If i < controlesVisibles.count Then
        On Error Resume Next
        controlesVisibles(i + 1).SetFocus
        On Error GoTo 0
    End If
End Sub

Private Sub cbxAlumno_Enter()
    Call SincronizarCedulaAlumno
End Sub
Private Sub cbxCAlumno_Enter()
    Call SincronizarAlumnoCedula
End Sub

Private Sub cbxBanco_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim match As String
        match = BuscarMejorCoincidencia(cbxBanco.Text, cbxBanco)
        If match <> "" Then
            cbxBanco.Text = match
        End If
        KeyCode = 0
    End If
End Sub

Private Sub cbxBCodigo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim match As String
        match = BuscarMejorCoincidencia(cbxBCodigo.Text, cbxBCodigo)
        If match <> "" Then
            cbxBCodigo.Text = match
            SincronizarBancoPorCodigo match
        End If
        KeyCode = 0
    End If
End Sub

Private Sub cbxAlumno_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim ws As Worksheet
        Dim nombreBuscar As String
        Dim mejorCoincidencia As String
        Dim partes() As String
        Dim nombre As String, apellido As String
        Dim fila As Long
        Dim ultimaFila As Long
        Dim cedulasEncontradas As Collection
        Dim item As Variant

        Set ws = ThisWorkbook.Sheets("Datos")
        nombreBuscar = LCase(Trim(cbxAlumno.Text))
        ultimaFila = ws.Cells(ws.Rows.count, "A").End(xlUp).row
        Set cedulasEncontradas = New Collection

        ' Buscar mejor coincidencia
        For fila = 2 To ultimaFila
            Dim nombreCompleto As String
            nombreCompleto = Trim(ws.Cells(fila, 3).Value & " " & ws.Cells(fila, 4).Value)
            If LCase(nombreCompleto) Like "*" & nombreBuscar & "*" Then
                mejorCoincidencia = nombreCompleto
                cbxAlumno.Text = mejorCoincidencia
                Exit For
            End If
        Next fila

        ' Obtener nombre y apellido exactos
        partes = Split(cbxAlumno.Text, " ")
        If UBound(partes) >= 1 Then
            nombre = partes(0)
            apellido = partes(1)

            cbxCAlumno.Clear

            ' Buscar todas las cédulas coincidentes
            For fila = 2 To ultimaFila
                If Trim(ws.Cells(fila, 3).Value) = nombre And Trim(ws.Cells(fila, 4).Value) = apellido Then
                    cedulasEncontradas.Add Trim(ws.Cells(fila, 5).Value)
                End If
            Next fila

            ' Llenar ComboBox con cédulas encontradas
            For Each item In cedulasEncontradas
                cbxCAlumno.AddItem item
            Next item

            If cedulasEncontradas.count > 0 Then
                cbxCAlumno.Text = cedulasEncontradas(1)
                cbxCedulaD.Text = cedulasEncontradas(1)
            Else
                cbxCAlumno.Text = ""
                cbxCedulaD.Text = ""
            End If
        End If

        KeyCode = 0
    End If
End Sub

Private Sub cbxCAlumno_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim ws As Worksheet
        Dim cedulaBuscar As String
        Dim mejorCoincidencia As String
        Dim fila As Long
        Dim ultimaFila As Long
        Dim nombresEncontrados As Collection
        Dim item As Variant

        Set ws = ThisWorkbook.Sheets("Datos")
        cedulaBuscar = UCase(Trim(cbxCAlumno.Text))
        ultimaFila = ws.Cells(ws.Rows.count, "A").End(xlUp).row
        Set nombresEncontrados = New Collection

        ' Autocompletar cédula
        For fila = 2 To ultimaFila
            If UCase(Trim(ws.Cells(fila, 5).Value)) Like "*" & cedulaBuscar & "*" Then
                mejorCoincidencia = Trim(ws.Cells(fila, 5).Value)
                cbxCAlumno.Text = mejorCoincidencia
                Exit For
            End If
        Next fila

        ' Buscar nombres/apellidos correspondientes
        cbxAlumno.Clear
        For fila = 2 To ultimaFila
            If UCase(Trim(ws.Cells(fila, 5).Value)) = cbxCAlumno.Text Then
                nombresEncontrados.Add Trim(ws.Cells(fila, 3).Value & " " & ws.Cells(fila, 4).Value)
            End If
        Next fila

        ' Llenar cbxAlumno
        For Each item In nombresEncontrados
            cbxAlumno.AddItem item
        Next item

        If nombresEncontrados.count > 0 Then
            cbxAlumno.Text = nombresEncontrados(1)
            cbxCedulaD.Text = cbxCAlumno.Text
        Else
            cbxAlumno.Text = ""
            cbxCedulaD.Text = ""
        End If

        KeyCode = 0
    End If
End Sub

Private Sub btnGuardar_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call btnGuardar_Click
        KeyCode = 0
    End If
End Sub

Private Sub btnLimpiar_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call btnGuardar_Click
        KeyCode = 0
    End If
End Sub
-------------------------------------------------------------------------------