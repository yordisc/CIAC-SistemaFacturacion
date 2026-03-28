Public UsuarioActual As String
Public NivelAcceso As String
Dim ProgramadoParaGuardar As Boolean

Private Sub ReprotegerHojas()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        On Error Resume Next
        ws.Unprotect Password:="seguro"
        ws.Protect Password:="seguro", _
            UserInterfaceOnly:=True, _
            AllowSorting:=True, _
            AllowFiltering:=True, _
            AllowUsingPivotTables:=True
        On Error GoTo 0
    Next ws
End Sub

Private Sub ProgramarGuardado()
    If Not ProgramadoParaGuardar Then
        ProgramadoParaGuardar = True
        GuardarAutomaticamente
    End If
End Sub

Private Sub GuardarAutomaticamente()
    On Error Resume Next
    ThisWorkbook.Save
    ' Reprogramar la próxima vez en 3 minutos
    Application.OnTime Now + TimeValue("00:03:00"), "GuardarAutomaticamente"
End Sub

Private Sub DetenerGuardadoAutomatico()
    On Error Resume Next
    Application.OnTime EarliestTime:=Now + TimeValue("00:03:00"), Procedure:="GuardarAutomaticamente", Schedule:=False
    ProgramadoParaGuardar = False
End Sub

Sub IniciarSesion()
    Dim usr As String
    Dim pwd As String
    Dim intentos As Integer
    Dim ws As Worksheet
    intentos = 0

    ' Activar y mostrar la hoja "Bienvenido"
    Sheets("Bienvenido").visible = xlSheetVisible
    Sheets("Bienvenido").Activate

    ' Ocultar todas las demás hojas
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Bienvenido" Then
            ws.visible = xlSheetVeryHidden
        End If
    Next ws

    ' Asegurar que la cinta esté oculta
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"

    ' Iniciar proceso de login de usuario
    Do
        usr = InputBox("Ingrese su usuario:", "Inicio de Sesión")
        If usr = vbNullString Then
            ThisWorkbook.Close SaveChanges:=False
        End If

        If ValidarUsuario(usr) Then
            Exit Do
        Else
            MsgBox "Usuario inválido.", vbExclamation
            intentos = intentos + 1
            If intentos >= 6 Then
                ThisWorkbook.Close SaveChanges:=False
            End If
        End If
    Loop

    ' Reiniciar contador para contraseña
    intentos = 0

    ' Ingresar contraseña
    Do
        pwd = InputBox("Ingrese su contraseña:", "Confirmar Contraseña") ' **(Aquí luego pondremos un UserForm si quieres más seguridad)**
        If pwd = vbNullString Then
            ThisWorkbook.Close SaveChanges:=False
        End If

        If ValidarPassword(usr, pwd) Then
            Exit Do
        Else
            MsgBox "Contraseña incorrecta.", vbExclamation
            intentos = intentos + 1
            If intentos >= 6 Then
                ThisWorkbook.Close SaveChanges:=False
            End If
        End If
    Loop

    ' Guardar el usuario logueado
    UsuarioActual = usr
    NivelAcceso = usr

    ' Configurar privilegios según el usuario
    Call ConfigurarAcceso

    ' Luego del login, si es usuario "01" o "02" ocultamos "Bienvenido"
    If NivelAcceso = "01" Or NivelAcceso = "02" Then
        Sheets("Bienvenido").visible = xlSheetVeryHidden
    End If

    ' Iniciar guardado automático
    Call ProgramarGuardado

    Application.ScreenUpdating = True
End Sub

Private Function ValidarUsuario(usr As String) As Boolean
    Select Case usr
        Case "00", "01", "02"
            ValidarUsuario = True
        Case Else
            ValidarUsuario = False
    End Select
End Function

Private Function ValidarPassword(usr As String, pwd As String) As Boolean
    Select Case usr
        Case "00": ValidarPassword = (pwd = "wer2,rwfwt34r")
        Case "01": ValidarPassword = (pwd = "clave01")
        Case "02": ValidarPassword = (pwd = "clave02")
        Case Else: ValidarPassword = False
    End Select
End Function

Private Sub ConfigurarAcceso()
    Dim ws As Worksheet

    ' Siempre mostrar Dashboard inicialmente
    Sheets("Dashboard").visible = xlSheetVisible
    Sheets("Dashboard").Activate

    ' Ocultar todo al iniciar
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Dashboard" Then
            ws.visible = xlSheetVeryHidden
        End If
    Next ws

    Select Case NivelAcceso
        Case "00" ' Técnico (acceso completo)
            ' Mostrar todas las hojas
            For Each ws In ThisWorkbook.Sheets
                ws.visible = xlSheetVisible
                ws.Unprotect Password:="seguro"
            Next ws
            ' Mostrar controles adicionales si es necesario
            On Error Resume Next
            With Sheets("Dashboard")
                .OLEObjects("btnSincronizarFacturas").visible = True
                .OLEObjects("btnRestaurarFacturas").visible = True
            End With
            On Error GoTo 0
            
            ' MANTENER la cinta visible
            Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", True)"

        Case "01" ' Administrador
            ' Mostrar solo las hojas permitidas
            Sheets("Dashboard").visible = xlSheetVisible
            Sheets("Buscar").visible = xlSheetVisible
            Sheets("Factura").visible = xlSheetVisible

            ' Proteger hojas
            Call ProtegerHojasAdmin

            ' Ocultar cinta
            Application.OnTime Now + TimeValue("00:00:05"), "OcultarCinta"
            Sheets("Dashboard").Activate
            
            Call PrepararTablasDinamicas
            Call HabilitarEventosDashboard
            Application.OnTime Now + TimeValue("00:00:01"), "HabilitarEventosDashboard"
            
            ' Habilitar botón para ActualizarCampoCalculado si existe
            On Error Resume Next
            With Sheets("Dashboard")
                .OLEObjects("btnActualizarTasa").visible = True
                .OLEObjects("btnActualizarTasa").Enabled = True
            End With
            On Error GoTo 0

        Case "02" ' Asistente
            ' Mostrar solo Dashboard y Factura
            Sheets("Dashboard").visible = xlSheetVisible
            Sheets("Factura").visible = xlSheetVisible

            ' Ocultar botones en Dashboard
            On Error Resume Next
            With Sheets("Dashboard")
                .OLEObjects("btnSincronizarFacturas").visible = False
                .OLEObjects("btnRestaurarFacturas").visible = False
                .OLEObjects("btnActualizarTasa").visible = True
                .OLEObjects("btnActualizarTasa").Enabled = True
            End With
            On Error GoTo 0

            ' Proteger hojas
            Call ProtegerHojasAsistente

            ' Ocultar cinta
            Application.OnTime Now + TimeValue("00:00:05"), "OcultarCinta"
            Sheets("Dashboard").Activate
            
            Call PrepararTablasDinamicas
            Call HabilitarEventosDashboard
            Application.OnTime Now + TimeValue("00:00:01"), "HabilitarEventosDashboard"
    End Select
End Sub

Private Sub OcultarCinta()
    On Error Resume Next
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
End Sub

Private Sub ProtegerHojasAdmin()
    Dim ws As Worksheet
    Dim rngEditable As Range
    
    For Each ws In ThisWorkbook.Sheets
        ws.Unprotect Password:="seguro"
        
        Select Case ws.Name
            Case "Buscar"
                Set rngEditable = ws.Range("A2:P1048576")
                ws.Cells.Locked = True
                rngEditable.Locked = False
                ws.Protect Password:="seguro", _
                    UserInterfaceOnly:=True, _
                    AllowSorting:=True, _
                    AllowFiltering:=True, _
                    AllowUsingPivotTables:=True, _
                    AllowFormattingColumns:=True, _
                    AllowInsertingColumns:=True, _
                    AllowDeletingColumns:=True
                    
            Case "Facturas", "Extras"
                Set rngEditable = ws.Range("A2:S1048576")
                ws.Cells.Locked = True
                rngEditable.Locked = False
                ws.Protect Password:="seguro", _
                    UserInterfaceOnly:=True, _
                    AllowFormattingColumns:=True, _
                    AllowInsertingColumns:=True, _
                    AllowDeletingColumns:=True, _
                    AllowUsingPivotTables:=True
                    
            Case "Datos", "R1", "R2", "Log"
                ' Configuración estándar para estas hojas
                Set rngEditable = ws.UsedRange
                ws.Cells.Locked = True
                rngEditable.Locked = False
                ws.Protect Password:="seguro", UserInterfaceOnly:=True
                
            Case "Extras"
                ' Desproteger completamente para tablas dinámicas
                ws.Unprotect Password:="seguro"
                
                ' Configurar tablas dinámicas para permitir cambios
                For Each pt In ws.PivotTables
                    pt.EnableDataValueEditing = True
                    pt.EnableDrilldown = True
                    pt.EnableFieldDialog = True
                    pt.EnableFieldList = True
                    pt.EnableWizard = True
                Next pt
                
                ' Proteger la hoja con parámetros especiales
                ws.Protect Password:="seguro", _
                    UserInterfaceOnly:=True, _
                    DrawingObjects:=True, _
                    Contents:=True, _
                    AllowFormattingCells:=True, _
                    AllowFormattingColumns:=True, _
                    AllowInsertingColumns:=True, _
                    AllowDeletingColumns:=True, _
                    AllowUsingPivotTables:=True, _
                    AllowFiltering:=True
                
                ws.EnableSelection = xlUnlockedCells
                
                
            Case "Dashboard"
                ' Configuración especial para Dashboard
                ws.Unprotect Password:="seguro"
                ws.Range("B6").Locked = False
                ws.Range("E3:E5").Locked = False
                ws.Protect Password:="seguro", _
                    UserInterfaceOnly:=True, _
                    DrawingObjects:=False, _
                    Contents:=True, _
                    AllowFormattingCells:=True, _
                    AllowUsingPivotTables:=True, _
                    AllowFiltering:=True
                ws.EnableSelection = xlUnlockedCells
                
            Case Else
                ws.Protect Password:="seguro", UserInterfaceOnly:=True
        End Select
    Next ws
End Sub

Private Sub ProtegerHojasAsistente()
    Dim ws As Worksheet
    Dim rngEditable As Range
    Const pwd As String = "seguro"
    
    Application.ScreenUpdating = False
    
    For Each ws In ThisWorkbook.Sheets
        On Error Resume Next
        ws.Unprotect Password:=pwd
        On Error GoTo 0
        
        Select Case ws.Name
            Case "Factura"
                ' Proteger Factura y permitir edición de celdas desbloqueadas
                With ws
                    .Cells.Locked = True
                    .Protect Password:=pwd, _
                             UserInterfaceOnly:=True, _
                             DrawingObjects:=True, _
                             Contents:=True, _
                             AllowFormattingCells:=True, _
                             AllowUsingPivotTables:=True, _
                             AllowFiltering:=True
                    .EnableSelection = xlUnlockedCells
                End With
                
            Case "Dashboard"
                With ws
                    ' 1) Desproteger
                    .Unprotect Password:=pwd
                    
                    ' 2) Bloquear todas las celdas
                    .Cells.Locked = True

                    ' 3) Configurar textos fijos y bloquear E3:E5
                    .Range("E3").Value = "COMBUSTIBLE DESPACHADO"
                    .Range("E4").Value = "AERONAVES MÁS TRIPULADAS"
                    .Range("E5").Value = "DISTRIBUCIÓN TIPO DE FACTURA"
                    .Range("E3:E5").Locked = True

                    ' 4) Desbloquear B6 explícitamente
                    .Range("B6").Locked = False

                    ' 5) Eliminar cualquier rango previo “EditB6” y crear uno nuevo
                    On Error Resume Next
                    .Protection.AllowEditRanges("EditB6").Delete
                    On Error GoTo 0
                    .Protection.AllowEditRanges.Add _
                        Title:="EditB6", _
                        Range:=.Range("B6"), _
                        Password:=""

                    ' 6) Proteger hoja
                    .Protect Password:=pwd, _
                             UserInterfaceOnly:=True, _
                             DrawingObjects:=False, _
                             Contents:=True, _
                             AllowFormattingCells:=True, _
                             AllowFiltering:=True, _
                             AllowUsingPivotTables:=True

                    ' 7) Solo permitir selección de celdas desbloqueadas
                    .EnableSelection = xlUnlockedCells

              End With
                
            Case "Extras"
                With ws
                    .Unprotect Password:=pwd
                    ' Permitir edición en todas las tablas dinámicas
                    Dim pt As PivotTable
                    For Each pt In .PivotTables
                        pt.EnableDataValueEditing = True
                        pt.EnableDrilldown = True
                        pt.EnableFieldDialog = True
                        pt.EnableFieldList = True
                        pt.EnableWizard = True
                    Next pt
                    .Cells.Locked = True
                    .Protect Password:=pwd, _
                             UserInterfaceOnly:=True, _
                             DrawingObjects:=True, _
                             Contents:=True, _
                             AllowFormattingCells:=True, _
                             AllowFormattingColumns:=True, _
                             AllowInsertingColumns:=True, _
                             AllowDeletingColumns:=True, _
                             AllowUsingPivotTables:=True, _
                             AllowFiltering:=True
                    .EnableSelection = xlUnlockedCells
                End With
                
            Case "Facturas"
                With ws
                    .Cells.Locked = True
                    Set rngEditable = .Range("A2:S1048576")
                    rngEditable.Locked = False
                    .Protect Password:=pwd, _
                             UserInterfaceOnly:=True, _
                             AllowFormattingColumns:=True, _
                             AllowInsertingColumns:=True, _
                             AllowDeletingColumns:=True, _
                             AllowUsingPivotTables:=True
                    .EnableSelection = xlUnlockedCells
                End With
                
            Case Else
                ' Para cualquier otra hoja, protección estándar
                ws.Cells.Locked = True
                ws.Protect Password:=pwd, UserInterfaceOnly:=True
        End Select
    Next ws
    
    Application.ScreenUpdating = True
End Sub

Private Sub DesbloquearTablaBusqueda()
    With Sheets("Buscar").Range("A2:P1048576")
        .Locked = False
    End With
End Sub

Public Sub HabilitarEventosDashboard()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ws.Unprotect Password:="seguro"
    ws.Protect Password:="seguro", _
        UserInterfaceOnly:=True, _
        DrawingObjects:=False, _
        Contents:=True, _
        AllowFormattingCells:=True, _
        AllowUsingPivotTables:=True, _
        AllowFiltering:=True
    ws.EnableSelection = xlUnlockedCells
    
    ' Habilitar/mostrar botones según nivel de acceso
    Dim btn As OLEObject
    For Each btn In ws.OLEObjects
        If TypeName(btn.Object) = "CommandButton" Then
            Select Case NivelAcceso
                Case "02"
                    ' Para el usuario 02: ocultar siempre estos dos
                    If btn.Name = "btnSincronizarFacturas" Or btn.Name = "btnRestaurarFacturas" Then
                        btn.visible = False
                    Else
                        btn.visible = True
                        btn.Object.Enabled = True
                    End If
                Case Else
                    ' Para otros usuarios: mostrarlos todos
                    btn.visible = True
                    btn.Object.Enabled = True
            End Select
        End If
    Next btn
End Sub

Public Sub PrepararTablasDinamicas()
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Extras")
    ws.Unprotect Password:="seguro"
    
    For Each pt In ws.PivotTables
        With pt
            .EnableDataValueEditing = True
            .EnableDrilldown = True
            .EnableFieldDialog = True
            .EnableFieldList = True
            .EnableWizard = True
            .PivotCache.MissingItemsLimit = xlMissingItemsNone
        End With
    Next pt
    
    ' Vuelve a proteger la hoja
    ws.Protect Password:="seguro", _
        UserInterfaceOnly:=True, _
        DrawingObjects:=True, _
        Contents:=True, _
        AllowFormattingCells:=True, _
        AllowFormattingColumns:=True, _
        AllowInsertingColumns:=True, _
        AllowDeletingColumns:=True, _
        AllowUsingPivotTables:=True, _
        AllowFiltering:=True
    
    On Error GoTo 0
End Sub
-------------------------------------------------------------------------------