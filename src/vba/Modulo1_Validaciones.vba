Option Explicit

' Validar nombre y apellido
Public Function EsValidoNombreApellido(nombreCompleto As String) As Boolean
    Dim partes() As String
    partes = Split(Trim(nombreCompleto))
    EsValidoNombreApellido = (UBound(partes) = 1)
End Function

' Validar cédula
Public Function EsCedulaValida(ced As String) As Boolean
    EsCedulaValida = False
    If ced Like "[vVeE][0-9]*" Then
        EsCedulaValida = True
    End If
End Function

Public Function EsMontoValido(monto As String) As Boolean
    Dim valoresPermitidos As Variant
    valoresPermitidos = Array("DIVISAS", "BOLIVARES")
    Dim valor As Variant
    For Each valor In valoresPermitidos
        If UCase(monto) = valor Then
            EsMontoValido = True
            Exit Function
        End If
    Next
    EsMontoValido = False
End Function

Public Function EsMetodoPagoValido(mpago As String) As Boolean
    Dim valoresPermitidos As Variant
    valoresPermitidos = Array("EFECTIVO", "PAGOMOVIL")
    Dim valor As Variant
    For Each valor In valoresPermitidos
        If UCase(mpago) = valor Then
            EsMetodoPagoValido = True
            Exit Function
        End If
    Next
    EsMetodoPagoValido = False
End Function

' Cargar datos en ComboBox
Public Sub CargarEnComboBox(wsName As String, rango As String, cbx As Object)
    Dim ws As Worksheet
    Dim rng As Range
    Dim celda As Range
    
    Set ws = ThisWorkbook.Worksheets(wsName)
    Set rng = ws.Range(rango)
    
    cbx.Clear
    For Each celda In rng
        If celda.Value <> "" Then
            On Error Resume Next
            cbx.AddItem celda.Value
            On Error GoTo 0
        End If
    Next celda
End Sub

Sub VerNombresControles()
    Dim obj As OLEObject
    For Each obj In Sheets("Factura").OLEObjects
        Debug.Print obj.Name
    Next obj
End Sub

Public Sub MostrarCamposPagoMovil(visible As Boolean)
    With Sheets("Factura")
        .OLEObjects("cbxBanco").visible = visible
        .OLEObjects("cbxBCodigo").visible = visible
        .OLEObjects("cbxCedulaD").visible = visible
        .OLEObjects("cbxNumOperacion").visible = visible
        .OLEObjects("cbxNumTlOrigen").visible = visible
    End With
End Sub

Sub QuickSort(arr() As String, first As Long, last As Long)
    Dim low As Long, high As Long
    Dim pivot As String, temp As String

    low = first
    high = last
    pivot = arr((first + last) \ 2)

    Do While low <= high
        Do While arr(low) < pivot
            low = low + 1
        Loop
        Do While arr(high) > pivot
            high = high - 1
        Loop
        If low <= high Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop

    If first < high Then QuickSort arr, first, high
    If low < last Then QuickSort arr, low, last
End Sub
-------------------------------------------------------------------------------