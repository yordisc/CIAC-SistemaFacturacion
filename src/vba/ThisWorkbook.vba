'-----------------------------------------------------------
' Módulo ThisWorkbook - Implementación completa
'-----------------------------------------------------------

Private Sub Workbook_Open()
    On Error GoTo ErrorHandler
    Dim startTime As Double
    startTime = Timer
    
    ' Configuración inicial de aplicación
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    ' Inicializar sistema de logs
    SetupLogSheet
    LogDebug "Inicio de sesión iniciado", Me.Name
    
    ' Mostrar Bienvenido
    Sheets("Bienvenido").visible = xlSheetVisible
    Sheets("Bienvenido").Activate
    Application.ScreenUpdating = True ' <- Aquí actívalo para que actualice pantalla y se vea
    
    ' Proteger todo antes de iniciar sesión
    Call ProtegerTodoInicialmente
    LogDebug "Protección inicial aplicada", Me.Name
    
    ' Procedimiento de login
    Call IniciarSesion
    LogDebug "Proceso de inicio de sesión completado", Me.Name
    
    ' Configuración final
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    ' Registrar tiempo de carga
    LogDebug "Workbook_Open completado en " & Format(Timer - startTime, "0.00") & " segundos", Me.Name
    
    Exit Sub
    
ErrorHandler:
    ' Registrar error y continuar
    GlobalErrorHandler Err.Number, Err.Description, "Workbook_Open"
    Resume Next
End Sub

'-----------------------------------------------------------
' Funciones de Logging
'-----------------------------------------------------------

' Configura la hoja Log si no existe
Private Sub SetupLogSheet()
    On Error Resume Next
    Dim wsLog As Worksheet
    Set wsLog = Me.Sheets("Log")
    
    If wsLog Is Nothing Then
        Set wsLog = Me.Sheets.Add(After:=Me.Sheets(Me.Sheets.count))
        wsLog.Name = "Log"
        With wsLog
            .Range("A1:C1").Value = Array("Fecha", "Archivo", "Descripción")
            .Columns("A:C").AutoFit
            .Tab.color = RGB(255, 200, 200) ' Color distintivo
        End With
        LogDebug "Hoja de Log creada", Me.Name
    End If
End Sub

' Función pública para registrar logs desde cualquier lugar
Public Sub LogDebug(ByVal descripcion As String, Optional ByVal archivo As String = "")
    On Error Resume Next
    
    Dim wsLog As Worksheet
    Dim nuevaFila As Long
    Const MAX_REGISTROS As Long = 50000 ' Conservar hasta 50,000 registros
    
    Set wsLog = Me.Sheets("Log")
    If wsLog Is Nothing Then SetupLogSheet: Set wsLog = Me.Sheets("Log")
    If wsLog Is Nothing Then Exit Sub
    
    ' Verificar cantidad de registros
    nuevaFila = wsLog.Cells(wsLog.Rows.count, "A").End(xlUp).row + 1
    
    ' Rotar logs si se excede el máximo
    If nuevaFila > MAX_REGISTROS + 1 Then
        RotarLogs MAX_REGISTROS
        nuevaFila = wsLog.Cells(wsLog.Rows.count, "A").End(xlUp).row + 1
    End If
    
    ' Registrar el nuevo evento
    With wsLog
        .Cells(nuevaFila, 1).Value = Now
        .Cells(nuevaFila, 2).Value = IIf(archivo = "", Me.Name, archivo)
        .Cells(nuevaFila, 3).Value = descripcion
        If nuevaFila Mod 100 = 0 Then .Columns("A:C").AutoFit
    End With
End Sub

' Rotar logs conservando los N más recientes
Private Sub RotarLogs(ByVal registrosAConservar As Long)
    Dim wsLog As Worksheet
    Dim ultimaFila As Long
    Dim rangoMover As Range
    
    Set wsLog = Me.Sheets("Log")
    If wsLog Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ultimaFila = wsLog.Cells(wsLog.Rows.count, "A").End(xlUp).row
    
    ' Si hay suficientes registros para rotar
    If ultimaFila > registrosAConservar + 1 Then
        ' Conservar los últimos N registros (más los encabezados)
        Set rangoMover = wsLog.Range("A" & ultimaFila - registrosAConservar + 1 & ":C" & ultimaFila)
        
        ' Limpiar todo excepto encabezados
        wsLog.Range("A2:C" & ultimaFila).ClearContents
        
        ' Pegar los registros más recientes
        rangoMover.Copy wsLog.Range("A2")
        
        ' Registrar la rotación
        wsLog.Cells(registrosAConservar + 2, 1).Value = Now
        wsLog.Cells(registrosAConservar + 2, 2).Value = Me.Name
        wsLog.Cells(registrosAConservar + 2, 3).Value = "ROTACIÓN DE LOG: Se conservaron los últimos " & registrosAConservar & " registros"
    End If
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

' Manejador global de errores
Public Sub GlobalErrorHandler(ByVal errNumber As Long, ByVal errDescription As String, _
                            ByVal moduleName As String, Optional ByVal procedureName As String = "")
    Dim errorMsg As String
    errorMsg = "Error " & errNumber & " en " & moduleName
    If procedureName <> "" Then errorMsg = errorMsg & "." & procedureName
    errorMsg = errorMsg & ": " & errDescription
    
    LogDebug errorMsg, moduleName
    
    ' Mostrar mensaje al usuario (opcional)
    If errNumber <> 0 Then
        MsgBox "Ocurrió un error (" & errNumber & "): " & errDescription & vbCrLf & _
               "Se ha registrado en el log.", vbExclamation, "Error"
    End If
End Sub

'-----------------------------------------------------------
' Funciones de apoyo para el sistema existente
'-----------------------------------------------------------

Private Sub ProtegerTodoInicialmente()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        On Error Resume Next
        ws.Unprotect Password:="seguro"
        ws.visible = xlSheetVeryHidden
        ws.Cells.Locked = True
        ws.Protect Password:="seguro", UserInterfaceOnly:=True
        On Error GoTo 0
    Next ws
    
    ' Solo dejar visible Bienvenido
    Sheets("Bienvenido").visible = xlSheetVisible
    Sheets("Bienvenido").Activate
    
    ' Ocultar la cinta de opciones
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
End Sub

'-----------------------------------------------------------
' Eventos adicionales para logging automático
'-----------------------------------------------------------

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    LogDebug "Workbook cerrado por el usuario", Me.Name
    ' Limpiar logs antiguos (opcional)
    ' Call LimpiarLogsAntiguos(30) ' Conservar logs de 30 días
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    LogDebug "Hoja activada: " & Sh.Name, Me.Name
End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    LogDebug "Nueva hoja creada: " & Sh.Name, Me.Name
End Sub

-------------------------------------------------------------------------------