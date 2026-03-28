Option Explicit

' Macro principal para generar el dashboard completo
Public Sub GenerarDashboardCompleto()
    Call CrearPanelDeControl
    Call ActualizarDatos
    Call CrearTablasDinamicas
    Call CrearSlicerTemporal
    Call CrearSlicerTipoFactura
    Call CrearGraficosDashboard
End Sub

' Crear el panel de control
Private Sub CrearPanelDeControl()
    Dim ws As Worksheet
    Dim btn As Button
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Limpiar y preparar el área
    ws.Range("A1:H20").Clear
    ws.Range("A1:H20").Interior.color = RGB(255, 255, 255) ' fondo blanco
    
    ' Eliminar botón si ya existe
    On Error Resume Next
    For Each btn In ws.Buttons
        If Not Intersect(btn.TopLeftCell, ws.Range("D6")) Is Nothing Then
            btn.Delete
        End If
    Next btn
    On Error GoTo 0

    ' Título principal
    With ws.Range("A1:H1")
        .Merge
        .Value = " PANEL DE CONTROL"
        .Font.Bold = True
        .Font.Size = 16
        .Font.color = RGB(255, 255, 255)
        .Interior.color = RGB(0, 102, 204)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' --- CONTROLES IZQUIERDA ---
    ws.Range("A6").Value = "TASA Bs/USD:"
    
    With ws.Range("A3:A6")
        .Font.Bold = True
        .Interior.color = RGB(230, 230, 230)
    End With
    
    ' Configurar validaciones de datos
    ConfigurarValidaciones ws
    
    ' Tasa de cambio por defecto
    ws.Range("B6").Value = 40
    ws.Range("B6").NumberFormat = "0.00"
    
    ' Crear botón en D6
    Set btn = ws.Buttons.Add( _
        Left:=ws.Range("D6").Left, _
        Top:=ws.Range("D6").Top, _
        Width:=ws.Range("D6").Width + 20, _
        Height:=ws.Range("D6").Height + 4)
    
    With btn
        .OnAction = "ActualizarDatos"
        .Caption = "ACTUALIZAR TASA Bs/USD"
        .Font.Bold = True
        .Name = "btnActualizarTasa"
    End With
    
    ' --- CONTROLES DERECHA: VISTAS DE DASHBOARD ---
    ws.Range("D3").Value = "DIAGRAMA DE BARRAS VERTICAL:"
    ws.Range("D4").Value = "DIAGRAMA DE BARRAS HORIZONTAL:"
    ws.Range("D5").Value = "DIAGRAMA CIRCULAR:"
    
    With ws.Range("D3:D5")
        .Font.Bold = True
        .Interior.color = RGB(230, 230, 230)
    End With
    
    ' Configurar opciones para gráficos
    With ws.Range("E3").Validation
        .Delete
        .Add xlValidateList, , , "RECAUDACIÓN EN USD,COMBUSTIBLE DESPACHADO"
    End With
    
    With ws.Range("E4").Validation
        .Delete
        .Add xlValidateList, , , "INSTRUCTORES,AERONAVES MÁS TRIPULADAS"
    End With
    
    With ws.Range("E5").Validation
        .Delete
        .Add xlValidateList, , , "DISTRIBUCIÓN TIPO DE FACTURA,DESPACHADOR"
    End With
    
    ' --- RESUMEN DE MONTOS ---
    CrearResumenMontos ws
    
    ' Ajustes finales
    ws.Columns("A:H").AutoFit
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Panel de control creado correctamente", vbInformation
End Sub

' Configurar validaciones de datos para el panel de control
Private Sub ConfigurarValidaciones(ByVal ws As Worksheet)
    With ws.Range("B3").Validation
        .Delete
        .Add xlValidateList, , , "=Extras!A51:A58" ' TablaMomento
    End With
    
    With ws.Range("B4").Validation
        .Delete
        .Add xlValidateList, , , "=Extras!A45:A48" ' TablaFactura
    End With
End Sub

' Crear el área de resumen de montos
Private Sub CrearResumenMontos(ByVal ws As Worksheet)
    With ws
        .Range("G2").Value = "Resumen de Montos"
        .Range("G2").Font.Bold = True
        .Range("G2").Font.Size = 12
        
        .Range("F3").Value = "DIVISAS (USD)"
        .Range("F4").Value = "EQUIV: BS (USD)"
        .Range("F5").Value = "TOTAL GENERAL"
        
        .Range("F3:F5").Font.Bold = True
        .Range("F3:F5").Interior.color = RGB(245, 245, 245)
        
        ' Usar GETPIVOTDATA para obtener los totales
        .Range("G3").Formula = "=GETPIVOTDATA(""Recaudación USD"",Extras!$C$80,""TIPO DE PAGO"",""DIVISAS"")"
        .Range("G4").Formula = "=GETPIVOTDATA(""Recaudación USD"",Extras!$C$80,""TIPO DE PAGO"",""BOLIVARES"")"
        .Range("G5").Formula = "=GETPIVOTDATA(""Recaudación USD"",Extras!$C$80)"
        
        .Range("G3:G5").NumberFormat = "#,##0.00"
        .Range("G3:G5").Interior.color = RGB(242, 255, 229)
        
        ' Bordes para el resumen
        With .Range("F2:G6").Borders
            .LineStyle = xlContinuous
            .color = RGB(200, 200, 200)
            .Weight = xlThin
        End With
    End With
End Sub

PrivateSub CrearTablasDinamicas()
    Dim ptCache As PivotCache
    Dim pt As PivotTable
    Dim wsExtras As Worksheet
    Dim destCell As Range
    Dim startRow As Long
    Dim i As Integer

    ' Establecer hoja destino
    Set wsExtras = ThisWorkbook.Sheets("Extras")

    ' Crear el PivotCache desde la tabla DatosFactura
    Set ptCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:="DatosFactura", _
        Version:=xlPivotTableVersion15)

    ' Comenzar desde fila 80, columna C
    startRow = 80
    Set destCell = wsExtras.Cells(startRow, 3)

    ' Eliminar tablas dinámicas previas en el área
    For i = wsExtras.PivotTables.count To 1 Step -1
        wsExtras.PivotTables(i).TableRange2.Clear
    Next i

    ' ---------------------------
    ' 1. Recaudación en USD
    ' ---------------------------
    Set pt = ptCache.CreatePivotTable(TableDestination:=destCell, TableName:="RecaudacionUSD")
    With pt
        .PivotFields("FECHA_CORTA").Orientation = xlRowField
        .PivotFields("TIPO DE FACTURA").Orientation = xlPageField
        .PivotFields("TIPO DE PAGO").Orientation = xlColumnField
        .AddDataField .PivotFields("TOTAL_USD"), "Recaudación USD", xlSum
    End With
    Set destCell = destCell.Offset(25, 0)

    ' ---------------------------
    ' 2. Combustible Despachado
    ' ---------------------------
    Set pt = ptCache.CreatePivotTable(TableDestination:=destCell, TableName:="CombustibleDespachado")
    With pt
        .PivotFields("FECHA_CORTA").Orientation = xlRowField
        .PivotFields("TIPO DE FACTURA").Orientation = xlPageField
        .AddDataField .PivotFields("CANTIDAD DE COMBUSTIBLE DEL AERONAVE"), "Combustible", xlSum
    End With
    Set destCell = destCell.Offset(25, 0)

    ' ---------------------------
    ' 3. Top 5 Instructores
    ' ---------------------------
    Set pt = ptCache.CreatePivotTable(TableDestination:=destCell, TableName:="TopInstructores")
    With pt
        .PivotFields("FECHA_CORTA").Orientation = xlPageField
        .PivotFields("TIPO DE FACTURA").Orientation = xlPageField
        .PivotFields("INSTRUCTOR").Orientation = xlRowField
        .AddDataField .PivotFields("ID"), "Cantidad", xlCount
    End With
    Set destCell = destCell.Offset(25, 0)

    ' ---------------------------
    ' 4. Aeronaves más usadas
    ' ---------------------------
    Set pt = ptCache.CreatePivotTable(TableDestination:=destCell, TableName:="TopAeronaves")
    With pt
        .PivotFields("FECHA_CORTA").Orientation = xlPageField
        .PivotFields("TIPO DE FACTURA").Orientation = xlPageField
        .PivotFields("AERONAVE").Orientation = xlRowField
        .AddDataField .PivotFields("ID"), "Cantidad", xlCount
    End With
    Set destCell = destCell.Offset(25, 0)

    ' ---------------------------
    ' 5. Distribución por tipo de factura
    ' ---------------------------
    Set pt = ptCache.CreatePivotTable(TableDestination:=destCell, TableName:="DTipoFactura")
    With pt
        .PivotFields("FECHA_CORTA").Orientation = xlPageField
        .PivotFields("TIPO DE FACTURA").Orientation = xlRowField
        .AddDataField .PivotFields("ID"), "Cantidad", xlCount
    End With
    Set destCell = destCell.Offset(25, 0)

    ' ---------------------------
    ' 6. Distribución por despachador
    ' ---------------------------
    Set pt = ptCache.CreatePivotTable(TableDestination:=destCell, TableName:="DDespachadores")
    With pt
        .PivotFields("FECHA_CORTA").Orientation = xlPageField
        .PivotFields("TIPO DE FACTURA").Orientation = xlPageField
        .PivotFields("DESPACHADOR").Orientation = xlRowField
        .AddDataField .PivotFields("ID"), "Cantidad", xlCount
    End With

    MsgBox "¡Tablas dinámicas creadas exitosamente con formato de fecha corta!", vbInformation
End Sub

Private Sub CrearGraficosDashboard()
    Dim wsDash As Worksheet
    Dim wsExtras As Worksheet
    Dim pt As PivotTable
    Dim chartObj As ChartObject
    Dim tasa As Double
    Dim chartTop As Double, chartLeft As Double
    Dim txtBox As Shape
    Dim pf As PivotField
    Dim pi As PivotItem
    Dim i As Long
    
    ' Hojas
    Set wsExtras = Worksheets("Extras")
    Set wsDash = Worksheets("Dashboard")

    ' Obtener tasa del dólar
    tasa = wsDash.Range("B6").Value

    ' Limpiar gráficos existentes en Dashboard
    For Each chartObj In wsDash.ChartObjects
        chartObj.Delete
    Next chartObj
    For Each txtBox In wsDash.Shapes
        If txtBox.Type = msoTextBox Then txtBox.Delete
    Next txtBox

    chartTop = 50
    chartLeft = 20

    ' --- Gráfico 1: Recaudación Bs y USD ---
    Set pt = wsExtras.PivotTables("RecaudacionUSD")
    
    ' Ocultar "(blank)" en todos los campos posibles
    For Each pf In pt.PivotFields
        On Error Resume Next
        pf.PivotItems("(blank)").visible = False
        pf.PivotItems("(Blank)").visible = False
        pf.PivotItems("(vacío)").visible = False ' Para versiones en español
        On Error GoTo 0
    Next pf

    Set chartObj = wsDash.ChartObjects.Add( _
        Left:=chartLeft, Top:=chartTop, Width:=500, Height:=250)
        chartObj.Name = "Chart_Recaudacion"

    With chartObj.chart
        .SetSourceData Source:=pt.TableRange1
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Distribución de Recaudación (Bs / USD)"
        ' Configuración adicional para ocultar valores vacíos
        .DisplayBlanksAs = xlNotPlotted
    End With

    chartTop = chartTop + 270

    ' --- Gráfico 2: Combustible Despachado ---
    Set pt = wsExtras.PivotTables("CombustibleDespachado")
    
    ' Ocultar "(blank)" en todos los campos
    For Each pf In pt.PivotFields
        On Error Resume Next
        pf.PivotItems("(blank)").visible = False
        pf.PivotItems("(Blank)").visible = False
        On Error GoTo 0
    Next pf

    Set chartObj = wsDash.ChartObjects.Add( _
        Left:=chartLeft, Top:=chartTop, Width:=500, Height:=250)
        chartObj.Name = "Chart_Combustible"

    With chartObj.chart
        .SetSourceData Source:=pt.TableRange1
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Combustible Despachado por Fecha"
        .DisplayBlanksAs = xlNotPlotted
    End With

    chartTop = chartTop + 270

    ' --- Gráfico 3: Top 5 Instructores ---
    Set pt = wsExtras.PivotTables("TopInstructores")

    ' Ocultar "(blank)" en todos los campos
    For Each pf In pt.PivotFields
        On Error Resume Next
        pf.PivotItems("(blank)").visible = False
        pf.PivotItems("(Blank)").visible = False
        On Error GoTo 0
    Next pf

    Set chartObj = wsDash.ChartObjects.Add( _
        Left:=chartLeft, Top:=chartTop, Width:=500, Height:=250)
        chartObj.Name = "Chart_Instructores"


    With chartObj.chart
        .SetSourceData Source:=pt.TableRange1
        .ChartType = xlBarClustered
        .HasTitle = True
        .ChartTitle.Text = "Top 5 Instructores"
        .DisplayBlanksAs = xlNotPlotted
    End With

    chartTop = chartTop + 270

    ' --- Gráfico 4: Aeronaves más usadas ---
    Set pt = wsExtras.PivotTables("TopAeronaves")

    ' Ocultar "(blank)" en todos los campos
    For Each pf In pt.PivotFields
        On Error Resume Next
        pf.PivotItems("(blank)").visible = False
        pf.PivotItems("(Blank)").visible = False
        On Error GoTo 0
    Next pf

    Set chartObj = wsDash.ChartObjects.Add( _
        Left:=chartLeft, Top:=chartTop, Width:=500, Height:=250)
        chartObj.Name = "Chart_Aeronaves"


    With chartObj.chart
        .SetSourceData Source:=pt.TableRange1
        .ChartType = xlBarClustered
        .HasTitle = True
        .ChartTitle.Text = "Aeronaves más tripuladas"
        .DisplayBlanksAs = xlNotPlotted
    End With

    chartTop = chartTop + 270

    ' --- Gráfico 5: Tipo de Factura ---
    Set pt = wsExtras.PivotTables("DTipoFactura")

    ' Ocultar "(blank)" en todos los campos
    For Each pf In pt.PivotFields
        On Error Resume Next
        pf.PivotItems("(blank)").visible = False
        pf.PivotItems("(Blank)").visible = False
        On Error GoTo 0
    Next pf

    Set chartObj = wsDash.ChartObjects.Add( _
        Left:=chartLeft, Top:=chartTop, Width:=350, Height:=250)
        chartObj.Name = "Chart_TipoFactura"

    With chartObj.chart
        .SetSourceData Source:=pt.TableRange1
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = "Distribución de Tipo de Factura"
        .ApplyDataLabels
        .DisplayBlanksAs = xlNotPlotted
        ' Ocultar categorías con valor cero en gráficos de torta
        For i = 1 To .seriesCollection(1).Points.count
            If .seriesCollection(1).Points(i).HasDataLabel Then
                If .seriesCollection(1).Points(i).DataLabel.Text = "0" Or _
                   .seriesCollection(1).Points(i).DataLabel.Text = "(blank)" Then
                    .seriesCollection(1).Points(i).Format.Fill.visible = msoFalse
                    .seriesCollection(1).Points(i).HasDataLabel = False
                End If
            End If
        Next i
    End With

    chartTop = chartTop + 270

    ' --- Gráfico 6: Despachadores ---
    Set pt = wsExtras.PivotTables("DDespachadores")

    ' Ocultar "(blank)" en todos los campos
    For Each pf In pt.PivotFields
        On Error Resume Next
        pf.PivotItems("(blank)").visible = False
        pf.PivotItems("(Blank)").visible = False
        On Error GoTo 0
    Next pf

    Set chartObj = wsDash.ChartObjects.Add( _
        Left:=chartLeft, Top:=chartTop, Width:=350, Height:=250)
        chartObj.Name = "Chart_Despachadores"

    With chartObj.chart
        .SetSourceData Source:=pt.TableRange1
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = "Distribución por Despachador"
        .ApplyDataLabels
        .DisplayBlanksAs = xlNotPlotted
        ' Ocultar categorías con valor cero en gráficos de torta
        For i = 1 To .seriesCollection(1).Points.count
            If .seriesCollection(1).Points(i).HasDataLabel Then
                If .seriesCollection(1).Points(i).DataLabel.Text = "0" Or _
                   .seriesCollection(1).Points(i).DataLabel.Text = "(blank)" Then
                    .seriesCollection(1).Points(i).Format.Fill.visible = msoFalse
                    .seriesCollection(1).Points(i).HasDataLabel = False
                End If
            End If
        Next i
    End With

    MsgBox "Gráficos generados correctamente en la hoja 'Dashboard'.", vbInformation
End Sub

Private Sub LimpiarFilasVacias()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn
    Dim dataArray As Variant
    Dim i As Long, lastRow As Long
    Dim colIndex As Long
    Dim t As Double
    
    ' Iniciar cronómetro para medir performance
    t = Timer
    
    Set ws = ThisWorkbook.Worksheets("Facturas")
    
    ' Validar tabla
    On Error Resume Next
    Set tbl = ws.ListObjects("DatosFactura")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "Error: Tabla 'DatosFactura' no encontrada", vbCritical
        Exit Sub
    End If
    
    ' Buscar columna por nombre de cabecera
    On Error Resume Next
    Set col = tbl.ListColumns("TOTAL_USD")
    On Error GoTo 0
    
    If col Is Nothing Then
        MsgBox "Error: Columna 'TOTAL_USD' no encontrada en la tabla", vbCritical
        Exit Sub
    End If
    
    ' Configuración para máxima velocidad
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    ' Obtener datos en array
    dataArray = col.DataBodyRange.Value
    lastRow = UBound(dataArray, 1)
    
    ' Procesamiento en memoria
    For i = 1 To lastRow
        If dataArray(i, 1) = "0" Then
            dataArray(i, 1) = vbNullString
        End If
    Next i
    
    ' Escribir datos de vuelta
    col.DataBodyRange.Value = dataArray
    
    ' Restaurar configuración
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
    End With
End Sub

Sub ActualizarDatos()
    Call ActualizarCampoCalculado
    Call ActualizarTodasLasTablasDinamicas
    Call LimpiarFilasVacias
End Sub

Public Sub ActualizarCampoCalculado()
    Dim wsFacturas As Worksheet
    Dim tbl As ListObject
    Dim cambio As Double
    Dim datos As Variant
    Dim resultado() As Variant, fechas() As Variant
    Dim i As Long
    Dim colTipoPago As Long, colPago As Long, colUSDTotal As Long
    Dim colFechaRecibo As Long, colFechaCorta As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wsFacturas = ThisWorkbook.Sheets("Facturas")
    
    ' Desproteger la hoja temporalmente (con contraseña)
    wsFacturas.Unprotect Password:="seguro"
    
    Set tbl = wsFacturas.ListObjects("DatosFactura")

    ' Obtener tasa de cambio desde Dashboard
    cambio = ThisWorkbook.Sheets("Dashboard").Range("B6").Value

    ' Eliminar columnas existentes si aplica
    On Error Resume Next
    tbl.ListColumns("TOTAL_USD").Delete
    tbl.ListColumns("FECHA_CORTA").Delete
    On Error GoTo 0

    ' Agregar nueva columna para fecha corta (sin hora)
    With tbl.ListColumns.Add
        .Name = "FECHA_CORTA"
        colFechaCorta = .Index
    End With

    ' Agregar nueva columna para USD
    With tbl.ListColumns.Add
        .Name = "TOTAL_USD"
        colUSDTotal = .Index
    End With

    ' Obtener datos de la tabla
    datos = tbl.DataBodyRange.Value

    ' Localizar índices
    colTipoPago = tbl.ListColumns("TIPO DE PAGO").Index
    colPago = tbl.ListColumns("PAGO").Index
    colFechaRecibo = tbl.ListColumns("FECHA DE RECIBO").Index

    ' Redimensionar arrays
    ReDim resultado(1 To UBound(datos, 1), 1 To 1)
    ReDim fechas(1 To UBound(datos, 1), 1 To 1)

    ' Calcular valores en USD y fecha corta
    For i = 1 To UBound(datos, 1)
        ' Extraer solo la fecha (sin hora)
        If IsDate(datos(i, colFechaRecibo)) Then
            fechas(i, 1) = DateValue(datos(i, colFechaRecibo))
        Else
            fechas(i, 1) = ""
        End If
        
        ' Calcular valores en USD
        Select Case UCase(datos(i, colTipoPago))
            Case "BOLIVARES"
                If cambio <> 0 Then
                    resultado(i, 1) = datos(i, colPago) / cambio
                Else
                    resultado(i, 1) = 0
                End If
            Case "DIVISAS"
                resultado(i, 1) = datos(i, colPago)
            Case Else
                resultado(i, 1) = 0
        End Select
    Next i

    ' Asignar los resultados
    tbl.ListColumns("FECHA_CORTA").DataBodyRange.Value = fechas
    tbl.ListColumns("TOTAL_USD").DataBodyRange.Value = resultado

    ' Formatear la columna de fecha corta
    tbl.ListColumns("FECHA_CORTA").DataBodyRange.NumberFormat = "dd/mm/yyyy"

    ' Reproteger la hoja
    wsFacturas.Protect Password:="seguro", _
        UserInterfaceOnly:=True, _
        AllowSorting:=True, _
        AllowFiltering:=True, _
        AllowUsingPivotTables:=True

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Campos calculados actualizados correctamente.", vbInformation
End Sub

Private Sub ActualizarTodasLasTablasDinamicas()
    Dim pt As PivotTable
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets("Extras")

    For Each pt In ws.PivotTables
        pt.RefreshTable
    Next pt

    MsgBox "Las graficas han sido actualizadas.", vbInformation
End Sub

Private Sub CrearSlicerTemporal()
    ' Definir constantes necesarias para el ordenamiento
    Const xlAscending As Integer = 1
    Const xlSortAscending As Integer = 1
    Const xlSlicerCrossFilterShowItemsWithDataAtTop As Integer = 2
    Const xlSlicerNoCrossFilter As Integer = 1
    Const xlTabular As Integer = 1
    
    Dim wsDash As Worksheet
    Dim wsExtras As Worksheet
    Dim sc As SlicerCache
    Dim sl As Slicer
    Dim pt As PivotTable
    Dim rngPeriodos As Range
    Dim pf As PivotField
    Dim i As Long
    Dim fieldName As String
    Dim success As Boolean
    
    ' Configurar hojas
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsExtras = ThisWorkbook.Sheets("Extras")
    
    ' Obtener rango de periodos desde TablaMomento
    Set rngPeriodos = wsExtras.Range("A51:A57")
    
    ' Limpiar slicers existentes
    On Error Resume Next
    For Each sc In ThisWorkbook.SlicerCaches
        sc.Delete
    Next sc
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    
    ' Verificar si hay tablas dinámicas
    If wsExtras.PivotTables.count = 0 Then
        MsgBox "No hay tablas dinámicas para conectar el slicer", vbExclamation
        Exit Sub
    End If
    
    ' Nombre del campo a usar (ahora usamos FECHA_CORTA)
    fieldName = "FECHA_CORTA"
    success = False
    
    ' Intentar crear el SlicerCache
    For Each pt In wsExtras.PivotTables
        On Error Resume Next
        Set pf = pt.PivotFields(fieldName)
        If Not pf Is Nothing Then
            Set sc = ThisWorkbook.SlicerCaches.Add2(pt, fieldName)
            If Not sc Is Nothing Then
                success = True
                Exit For
            End If
        End If
        On Error GoTo 0
    Next pt
    
    ' Verificar si se creó correctamente
    If Not success Then
        MsgBox "No se pudo crear el SlicerCache. Verifique que el campo '" & fieldName & "' existe.", vbExclamation
        Exit Sub
    End If
    
    ' Configurar el SlicerCache
    sc.Name = "Slicer_FechaCorta"
    
    ' Conectar a todas las tablas dinámicas con este campo
    For Each pt In wsExtras.PivotTables
        On Error Resume Next
        Set pf = pt.PivotFields(fieldName)
        If Not pf Is Nothing Then
            ' Verificar si ya está conectada
            Dim yaConectada As Boolean
            yaConectada = False
            
            For i = 1 To sc.PivotTables.count
                If sc.PivotTables(i).Name = pt.Name Then
                    yaConectada = True
                    Exit For
                End If
            Next i
            
            If Not yaConectada Then
                sc.PivotTables.AddPivotTable pt
            End If
        End If
        On Error GoTo 0
    Next pt
    
    ' Configuración avanzada del Slicer (Excel 2010+)
    On Error Resume Next
    With sc
        ' 1. Configurar el tipo de filtrado cruzado
        .CrossFilterType = xlSlicerCrossFilterShowItemsWithDataAtTop
        
        ' 2. Mostrar todos los ítems
        .ShowAllItems = True
        
        ' 3. Ordenar los ítems (usando la constante definida)
        .SortItems = xlSortAscending
        
        ' 4. Para Excel 2013+ (versión 15.0+)
        If Val(Application.Version) >= 15 Then
            .SortUsingCustomLists = False
        End If
    End With
    On Error GoTo 0
    
    ' Asegurar el orden en las tablas dinámicas
    For i = 1 To sc.PivotTables.count
        On Error Resume Next
        With sc.PivotTables(i).PivotFields(fieldName)
            .AutoSort xlAscending, fieldName
            .LayoutForm = xlTabular
        End With
        On Error GoTo 0
    Next i
    
    ' Crear el slicer visible
    Set sl = sc.Slicers.Add(wsDash, , , "Periodo", 350, 30, 180, 200)
    
    ' Configurar apariencia
    With sl
        .Style = "SlicerStyleLight3"
        .NumberOfColumns = 1
        .Caption = "Seleccionar Periodo"
        .Width = 180
        .Height = 200
        .Name = "Slicer_Fechas"
    End With
    
    ' Agregar botones de periodos
    CrearBotonesPeriodos wsDash, rngPeriodos, sc
    
    ' Actualizar todo
    sc.ClearManualFilter
    For Each pt In sc.PivotTables
        pt.RefreshTable
    Next pt
    
    Application.ScreenUpdating = True
    
    MsgBox "Slicer creado exitosamente con ordenamiento ascendente.", vbInformation
End Sub

Private Sub CrearBotonesPeriodos(ws As Worksheet, rngPeriodos As Range, sc As SlicerCache)
    Dim btn As Shape
    Dim shp As Shape
    Dim i As Long
    Dim topPos As Long
    Dim btnWidth As Long
    Dim btnHeight As Long
    Dim leftPos As Long
    
    ' Configuración de botones
    btnWidth = 120
    btnHeight = 24
    leftPos = 350
    topPos = 240 ' Posición debajo del slicer
    
    ' Eliminar botones existentes (si los hay)
    On Error Resume Next
    For Each shp In ws.Shapes
        If shp.Top > 230 And shp.Left > 340 Then
            If shp.Type = msoFormControl Or shp.Type = msoAutoShape Then
                shp.Delete
            End If
        End If
    Next shp
    On Error GoTo 0
    
    ' Crear botones para cada periodo (usando Shapes para mejor formato)
    For i = 1 To rngPeriodos.Rows.count
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, btnWidth, btnHeight)
        
        With btn
            .TextFrame.Characters.Text = rngPeriodos.Cells(i, 1).Value
            .TextFrame.HorizontalAlignment = xlCenter
            .TextFrame.VerticalAlignment = xlCenter
            .TextFrame.Characters.Font.Bold = True
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.Characters.Font.color = RGB(255, 255, 255) ' Texto blanco
            .Fill.ForeColor.RGB = RGB(0, 112, 192) ' Fondo azul
            .Line.visible = msoFalse ' Sin borde
            
            ' Asignar macro
            .OnAction = "'AplicarFiltroPeriodo """ & rngPeriodos.Cells(i, 1).Value & """'"
            .Name = "btnPeriodo_" & Replace(rngPeriodos.Cells(i, 1).Value, " ", "_")
        End With
        
        topPos = topPos + btnHeight + 5
    Next i
    
    ' Agregar botón para limpiar filtros (con color rojo)
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, btnWidth, btnHeight)
    
    With btn
        .TextFrame.Characters.Text = "TODO"
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.Characters.Font.color = RGB(255, 255, 255) ' Texto blanco
        .Fill.ForeColor.RGB = RGB(192, 0, 0) ' Fondo rojo
        .Line.visible = msoFalse ' Sin borde
        
        ' Asignar macro
        .OnAction = "'LimpiarFiltrosPeriodo'"
        .Name = "btnPeriodo_TODO"
    End With
End Sub

Public Sub AplicarFiltroPeriodo(periodo As String)
    Dim wsDash As Worksheet
    Dim wsExtras As Worksheet
    Dim sc As SlicerCache
    Dim pt As PivotTable
    Dim fechaInicio As Date
    Dim fechaFin As Date
    Dim hoy As Date
    Dim pf As PivotField
    Dim item As PivotItem
    Dim fechaValida As Boolean
    
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsExtras = ThisWorkbook.Sheets("Extras")
    
    hoy = Date
    Application.ScreenUpdating = False
    fechaValida = True
    
    ' Obtener el slicer cache (ahora buscamos por FECHA_CORTA)
    On Error Resume Next
    Set sc = ThisWorkbook.SlicerCaches("Slicer_FechaCorta")
    If sc Is Nothing Then
        Set sc = ThisWorkbook.SlicerCaches("Slicer_Fechas")
    End If
    On Error GoTo 0
    
    If sc Is Nothing Then
        MsgBox "No se encontró el slicer de fechas. Por favor, ejecute primero 'CrearSlicerTemporal'.", vbExclamation
        Exit Sub
    End If
    
    ' Determinar rango de fechas según el periodo seleccionado
    Select Case UCase(periodo)
        Case "HOY"
            fechaInicio = hoy
            fechaFin = hoy
        Case "AYER"
            fechaInicio = hoy - 1
            fechaFin = hoy - 1
        Case "SEMANAL"
            fechaInicio = hoy - Weekday(hoy, vbSunday) + 1
            fechaFin = hoy
        Case "MENSUAL"
            fechaInicio = DateSerial(Year(hoy), Month(hoy), 1)
            fechaFin = hoy
        Case "TRIMESTRE"
            fechaInicio = DateSerial(Year(hoy), Int((Month(hoy) - 1) / 3) * 3 + 1, 1)
            fechaFin = hoy
        Case "SEMESTRE"
            If Month(hoy) <= 6 Then
                fechaInicio = DateSerial(Year(hoy), 1, 1)
            Else
                fechaInicio = DateSerial(Year(hoy), 7, 1)
            End If
            fechaFin = hoy
        Case "ANUAL"
            fechaInicio = DateSerial(Year(hoy), 1, 1)
            fechaFin = hoy
        Case Else
            fechaValida = False
            MsgBox "Periodo no reconocido: " & periodo, vbExclamation
    End Select
    
    If fechaValida Then
        ' Aplicar filtro a todas las tablas dinámicas conectadas
        For Each pt In sc.PivotTables
            On Error Resume Next
            Set pf = pt.PivotFields("FECHA_CORTA")
            On Error GoTo 0
            
            If Not pf Is Nothing Then
                ' Limpiar filtros previos
                pf.ClearAllFilters
                
                ' Aplicar filtro de fecha (solo fecha, sin hora)
                On Error Resume Next
                pf.PivotFilters.Add2 _
                    Type:=xlDateBetween, _
                    Value1:=fechaInicio, _
                    Value2:=fechaFin
                On Error GoTo 0
                    
                ' Actualizar la tabla
                pt.RefreshTable
            End If
        Next pt
        
        ' Actualizar el slicer para mostrar solo las fechas filtradas
        sc.ClearManualFilter
        
        MsgBox "Filtro aplicado: " & periodo & vbCrLf & _
               "Desde: " & Format(fechaInicio, "dd/mm/yyyy") & vbCrLf & _
               "Hasta: " & Format(fechaFin, "dd/mm/yyyy"), vbInformation
    End If
    
    Application.ScreenUpdating = True
End Sub

Public Sub LimpiarFiltrosPeriodo()
    Dim wsDash As Worksheet
    Dim wsExtras As Worksheet
    Dim sc As SlicerCache
    Dim pt As PivotTable
    Dim pf As PivotField
    
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsExtras = ThisWorkbook.Sheets("Extras")
    
    Application.ScreenUpdating = False
    
    ' Obtener el slicer cache (usando el nombre correcto)
    On Error Resume Next
    Set sc = ThisWorkbook.SlicerCaches("Slicer_Fechas")
    If sc Is Nothing Then
        Set sc = ThisWorkbook.SlicerCaches("Slicer_FechaCorta")
    End If
    On Error GoTo 0
    
    If sc Is Nothing Then
        MsgBox "No se encontró el slicer de fechas. Por favor, ejecute primero 'CrearSlicerTemporal'.", vbExclamation
        Exit Sub
    End If
    
    ' Limpiar filtros en todas las tablas dinámicas conectadas
    For Each pt In sc.PivotTables
        On Error Resume Next
        ' Usar el campo correcto "FECHA_CORTA" en lugar de "FECHA DE RECIBO"
        Set pf = pt.PivotFields("FECHA_CORTA")
        On Error GoTo 0
        
        If Not pf Is Nothing Then
            pf.ClearAllFilters
            pt.RefreshTable
        End If
    Next pt
    
    ' Limpiar filtro del slicer
    sc.ClearManualFilter
    
    ' Limpiar mensaje de filtro
    wsDash.Range("J1:J3").ClearContents
    
    Application.ScreenUpdating = True
    
    MsgBox "Todos los filtros de período han sido eliminados", vbInformation
End Sub

Private Sub CrearSlicerTipoFactura()
    Dim wsDash As Worksheet
    Dim wsExtras As Worksheet
    Dim sc As SlicerCache
    Dim sl As Slicer
    Dim pt As PivotTable
    Dim versionExcel As Double
    
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsExtras = ThisWorkbook.Sheets("Extras")
    
    ' Obtener versión de Excel
    versionExcel = Val(Application.Version)
    
    ' Verificar si hay tablas dinámicas
    If wsExtras.PivotTables.count = 0 Then
        MsgBox "No hay tablas dinámicas para conectar el slicer", vbExclamation
        Exit Sub
    End If
    
    ' Limpiar slicer existente si lo hay
    On Error Resume Next
    ThisWorkbook.SlicerCaches("Slicer_TipoFactura").Delete
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    
    ' Crear el SlicerCache usando la primera tabla dinámica
    Set pt = wsExtras.PivotTables(1)
    
    ' Usar método adecuado según versión de Excel
    If versionExcel >= 14 Then ' Excel 2010 o superior
        Set sc = ThisWorkbook.SlicerCaches.Add2(pt, "TIPO DE FACTURA")
    Else ' Excel 2007
        Set sc = ThisWorkbook.SlicerCaches.Add(pt, "TIPO DE FACTURA")
    End If
    
    sc.Name = "Slicer_TipoFactura"
    
    ' Conectar a todas las tablas dinámicas
    For Each pt In wsExtras.PivotTables
        On Error Resume Next
        sc.PivotTables.AddPivotTable pt
        On Error GoTo 0
    Next pt
    
    ' Configurar el slicer
    On Error Resume Next
    If versionExcel >= 14 Then
        sc.CrossFilterType = xlSlicerCrossFilterShowItemsWithDataAtTop
        sc.SortItems = xlAscending
    End If
    
    sc.ShowAllItems = True
    
    If versionExcel >= 15 Then
        sc.SortUsingCustomLists = False
    End If
    On Error GoTo 0
    
    ' Ordenar los items en las tablas dinámicas (compatible con todas las versiones)
    For Each pt In sc.PivotTables
        On Error Resume Next
        With pt.PivotFields("TIPO DE FACTURA")
            .AutoSort xlAscending, "TIPO DE FACTURA"
            .LayoutForm = xlTabular
        End With
        On Error GoTo 0
    Next pt
    
    ' Crear el slicer visible
    Set sl = sc.Slicers.Add(wsDash, , , "TipoFactura", 20, 240, 180, 150)
    
    With sl
        .Style = "SlicerStyleLight3"
        .NumberOfColumns = 1
        .Caption = "Tipo de Factura"
        .Width = 180
        .Height = 150
        .Name = "Slicer_TipoFactura"
    End With
    
    ' Actualizar todo
    sc.ClearManualFilter
    For Each pt In sc.PivotTables
        pt.RefreshTable
    Next pt
    
    Application.ScreenUpdating = True
    
    MsgBox "Slicer de tipo de factura creado exitosamente", vbInformation
End Sub

Private Sub CrearBotonesTipoFactura()
    Dim wsDash As Worksheet
    Dim rngTiposFactura As Range
    Dim btn As Shape
    Dim shp As Shape
    Dim i As Long
    Dim topPos As Long
    Dim btnWidth As Long
    Dim btnHeight As Long
    Dim leftPos As Long
    
    ' Configurar hoja y rango
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set rngTiposFactura = wsDash.Range("A46:A48") ' HONORARIO, COMBUSTIBLE, H&C
    
    ' Configuración de botones
    btnWidth = 120
    btnHeight = 24
    leftPos = 20  ' Posición a la izquierda del dashboard
    topPos = 240  ' Posición debajo de otros controles
    
    ' Eliminar botones existentes (si los hay)
    On Error Resume Next
    For Each shp In wsDash.Shapes
        If shp.Name Like "btnTipoFactura_*" Or shp.Name = "btnTipoFactura_TODO" Then
            shp.Delete
        End If
    Next shp
    On Error GoTo 0
    
    ' Crear título para los botones
    Set btn = wsDash.Shapes.AddShape(msoShapeRectangle, leftPos, topPos - 30, btnWidth, btnHeight)
    With btn
        .TextFrame.Characters.Text = "FILTRAR POR TIPO:"
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.Characters.Font.color = RGB(255, 255, 255) ' Texto blanco
        .Fill.ForeColor.RGB = RGB(0, 112, 192) ' Fondo azul
        .Line.visible = msoFalse ' Sin borde
        .Name = "lblTipoFactura"
    End With
    
    ' Crear botones para cada tipo de factura
    For i = 1 To rngTiposFactura.Rows.count
        Set btn = wsDash.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, btnWidth, btnHeight)
        
        With btn
            .TextFrame.Characters.Text = rngTiposFactura.Cells(i, 1).Value
            .TextFrame.HorizontalAlignment = xlCenter
            .TextFrame.VerticalAlignment = xlCenter
            .TextFrame.Characters.Font.Bold = True
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.Characters.Font.color = RGB(255, 255, 255) ' Texto blanco
            
            ' Asignar color según el tipo de factura
            Select Case UCase(rngTiposFactura.Cells(i, 1).Value)
                Case "HONORARIO"
                    .Fill.ForeColor.RGB = RGB(0, 176, 80) ' Verde
                Case "COMBUSTIBLE"
                    .Fill.ForeColor.RGB = RGB(255, 0, 0) ' Rojo
                Case "H&C"
                    .Fill.ForeColor.RGB = RGB(0, 176, 240) ' Azul claro
            End Select
            
            .Line.visible = msoFalse ' Sin borde
            
            ' Asignar macro
            .OnAction = "'AplicarFiltroTipoFactura """ & rngTiposFactura.Cells(i, 1).Value & """'"
            .Name = "btnTipoFactura_" & Replace(rngTiposFactura.Cells(i, 1).Value, " ", "_")
        End With
        
        topPos = topPos + btnHeight + 5
    Next i
    
    ' Agregar botón para mostrar TODO (todos los tipos)
    Set btn = wsDash.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, btnWidth, btnHeight)
    
    With btn
        .TextFrame.Characters.Text = "TODO"
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.Characters.Font.color = RGB(255, 255, 255) ' Texto blanco
        .Fill.ForeColor.RGB = RGB(128, 128, 128) ' Gris
        .Line.visible = msoFalse ' Sin borde
        
        ' Asignar macro
        .OnAction = "'LimpiarFiltrosTipoFactura'"
        .Name = "btnTipoFactura_TODO"
    End With
    
    ' Ocultar el ítem "(en blanco)" en todas las tablas dinámicas
    OcultarItemEnBlanco
    
    MsgBox "Botones de filtro por tipo de factura creados correctamente.", vbInformation
End Sub

Private Sub OcultarItemEnBlanco()
    Dim wsExtras As Worksheet
    Dim pt As PivotTable
    Dim pf As PivotField
    
    Set wsExtras = ThisWorkbook.Sheets("Extras")
    
    For Each pt In wsExtras.PivotTables
        On Error Resume Next
        Set pf = pt.PivotFields("TIPO DE FACTURA")
        If Not pf Is Nothing Then
            ' Intentar ocultar diferentes versiones de "en blanco"
            pf.PivotItems("(blank)").visible = False
            pf.PivotItems("(Blank)").visible = False
            pf.PivotItems("(vacío)").visible = False
            pf.PivotItems("(empty)").visible = False
        End If
        On Error GoTo 0
    Next pt
End Sub

Private Sub AplicarFiltroTipoFactura(tipoFactura As String)
    Dim wsExtras As Worksheet
    Dim wsDash As Worksheet
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim sc As SlicerCache
    
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsExtras = ThisWorkbook.Sheets("Extras")
    
    Application.ScreenUpdating = False
    
    ' Buscar slicer de tipo de factura si existe
    On Error Resume Next
    Set sc = ThisWorkbook.SlicerCaches("Slicer_TipoFactura")
    On Error GoTo 0
    
    If Not sc Is Nothing Then
        ' Usar slicer si existe
        sc.ClearManualFilter
        On Error Resume Next
        sc.PivotTables(1).PivotFields("TIPO DE FACTURA").CurrentPage = tipoFactura
        On Error GoTo 0
    Else
        ' Aplicar filtro directamente a las tablas dinámicas
        For Each pt In wsExtras.PivotTables
            On Error Resume Next
            Set pf = pt.PivotFields("TIPO DE FACTURA")
            If Not pf Is Nothing Then
                pf.ClearAllFilters
                pf.CurrentPage = tipoFactura
            End If
            On Error GoTo 0
        Next pt
    End If
    
    ' Actualizar mensaje en Dashboard
    wsDash.Range("K1").Value = "Filtro activo: " & tipoFactura
    
    ' Actualizar gráficos
    Call CrearGraficosDashboard
    
    Application.ScreenUpdating = True
    
    MsgBox "Filtro aplicado: " & tipoFactura, vbInformation
End Sub

Private Sub LimpiarFiltrosTipoFactura()
    Dim wsExtras As Worksheet
    Dim wsDash As Worksheet
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim sc As SlicerCache
    
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsExtras = ThisWorkbook.Sheets("Extras")
    
    Application.ScreenUpdating = False
    
    ' Buscar slicer de tipo de factura si existe
    On Error Resume Next
    Set sc = ThisWorkbook.SlicerCaches("Slicer_TipoFactura")
    On Error GoTo 0
    
    If Not sc Is Nothing Then
        ' Limpiar filtro del slicer
        sc.ClearManualFilter
    Else
        ' Limpiar filtros directamente en las tablas dinámicas
        For Each pt In wsExtras.PivotTables
            On Error Resume Next
            Set pf = pt.PivotFields("TIPO DE FACTURA")
            If Not pf Is Nothing Then
                pf.ClearAllFilters
            End If
            On Error GoTo 0
        Next pt
    End If
    
    ' Limpiar mensaje en Dashboard
    wsDash.Range("K1").ClearContents
    
    Application.ScreenUpdating = True
    
    MsgBox "Todos los filtros de tipo de factura han sido eliminados", vbInformation
End Sub
-------------------------------------------------------------------------------