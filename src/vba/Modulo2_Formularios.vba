Public Sub VerControlesHojaActiva()
    Dim obj As OLEObject
    For Each obj In ActiveSheet.OLEObjects
        Debug.Print obj.Name & " - " & TypeName(obj.Object)
    Next obj
End Sub

Public Sub MostrarTodosOLEObjects()
    Dim obj As OLEObject
    Dim mensaje As String
    
    mensaje = ""
    
    For Each obj In ActiveSheet.OLEObjects
        mensaje = mensaje & obj.Name & " - " & TypeName(obj.Object) & vbCrLf
    Next obj
    
    If mensaje = "" Then
        MsgBox "No hay OLEObjects en esta hoja."
    Else
        MsgBox mensaje
    End If
End Sub

Sub LimpiarCampos()

    Application.ScreenUpdating = False

    With Sheets("Factura")

        ' Limpiar valores de los ComboBoxes
        .OLEObjects("cbxFecha").Object.Text = ""
        .OLEObjects("cbxDespachador").Object.Text = ""
        .OLEObjects("cbxInstructor").Object.Text = ""
        .OLEObjects("cbxAlumno").Object.Text = ""
        .OLEObjects("cbxCAlumno").Object.Text = ""
        .OLEObjects("cbxAeronave").Object.Text = ""
        .OLEObjects("cbxLitros").Object.Text = ""
        .OLEObjects("cbxObservacion").Object.Text = ""
        .OLEObjects("cbxMonto").Object.Text = ""
        .OLEObjects("cbxMPago").Object.Text = ""
        .OLEObjects("cbxCantidad").Object.Text = ""
        .OLEObjects("cbxBanco").Object.Text = ""
        .OLEObjects("cbxBCodigo").Object.Text = ""
        .OLEObjects("cbxCedulaD").Object.Text = ""
        .OLEObjects("cbxNumOperacion").Object.Text = ""
        .OLEObjects("cbxNumTlOrigen").Object.Text = ""
        
    End With

    Application.ScreenUpdating = True


End Sub

Sub LimpiarTablaDatosFactura()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Facturas")
    
    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    If ultimaFila > 1 Then
        ws.Rows("2:" & ultimaFila).ClearContents
        MsgBox "La tabla 'DatosFactura' fue limpiada exitosamente.", vbInformation
    Else
        MsgBox "No hay datos para limpiar.", vbExclamation
    End If
End Sub

Sub LimpiarTablaDatosUsuarios()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Datos")
    
    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.count, "A").End(xlUp).row

    If ultimaFila > 2 Then
        ws.Range("A3:E" & ultimaFila).ClearContents ' Borra desde la fila 3 en adelante
        MsgBox "Se limpiaron los datos desde la fila 3 en adelante en 'DatosUsuarios'.", vbInformation
    Else
        MsgBox "No hay filas adicionales para limpiar en 'DatosUsuarios'.", vbExclamation
    End If
End Sub

Sub LimpiarRespaldos()
    Dim hojas As Variant
    Dim nombreHoja As Variant
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim ultimaColumna As Long

    hojas = Array("R1", "R2")

    For Each nombreHoja In hojas
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(nombreHoja)
        On Error GoTo 0

        If Not ws Is Nothing Then
            With ws
                ultimaFila = .Cells(.Rows.count, 1).End(xlUp).row
                ultimaColumna = .Cells(1, .Columns.count).End(xlToLeft).Column
                If ultimaFila > 1 Then
                    .Range(.Cells(2, 1), .Cells(ultimaFila, ultimaColumna)).ClearContents
                End If
            End With
        Else
            MsgBox "No se encontró la hoja '" & nombreHoja & "'", vbExclamation
        End If
    Next nombreHoja

    MsgBox "Datos de las hojas R1 y R2 han sido limpiados.", vbInformation
End Sub

Sub LimpiarHojaLog()
    On Error Resume Next
    
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets("Log")
    
    If wsLog Is Nothing Then
        MsgBox "No se encontró la hoja 'Log'.", vbExclamation
        Exit Sub
    End If
    
    ' Confirmar con el usuario
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Está seguro que desea limpiar toda la hoja Log?" & vbCrLf & _
                      "Se conservará la primera fila (encabezados).", _
                      vbQuestion + vbYesNo, "Confirmar limpieza")
    
    If respuesta = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Registrar acción antes de limpiar
    Dim ultimaFila As Long
    ultimaFila = wsLog.Cells(wsLog.Rows.count, "A").End(xlUp).row
    
    If ultimaFila > 1 Then
        ' Limpiar contenido
        wsLog.Range("A2:C" & ultimaFila).ClearContents
        
        ' Opcional: Limpiar formatos y validaciones de datos
        wsLog.Range("A2:C" & ultimaFila).ClearFormats
        wsLog.Range("A2:C" & ultimaFila).Validation.Delete
        
        ' Registrar la limpieza
        wsLog.Cells(2, 1).Value = Now
        wsLog.Cells(2, 2).Value = ThisWorkbook.Name
        wsLog.Cells(2, 3).Value = "LIMPIEZA MANUAL: Se borraron " & (ultimaFila - 1) & " registros"
        
        ' Autoajustar columnas
        wsLog.Columns("A:C").AutoFit
    Else
        MsgBox "La hoja Log ya está vacía.", vbInformation
    End If
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

Sub LimpiarUsuariosFactura()
    LimpiarTablaDatosFactura
    LimpiarTablaDatosUsuarios
End Sub

Sub LimpiarTodo()
    LimpiarUsuariosFactura
    LimpiarRespaldos
    LimpiarHojaLog
End Sub
-------------------------------------------------------------------------------