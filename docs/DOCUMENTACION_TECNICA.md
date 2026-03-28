# Documentación Técnica — Módulos VBA

Este documento describe la función y estructura de cada módulo VBA del sistema.

---

## ThisWorkbook.vba
**Tipo:** Módulo de libro (clase)

Contiene los eventos principales del libro de trabajo y el sistema de logging.

### Funciones/Procedimientos clave

| Nombre | Tipo | Descripción |
|---|---|---|
| `Workbook_Open` | Evento | Punto de entrada del sistema. Inicializa logs, protege hojas, muestra pantalla de bienvenida y lanza el proceso de login. |
| `LogDebug` | Público | Registra cualquier evento con timestamp en la hoja `Log`. Admite rotación automática al superar 50,000 registros. |
| `GlobalErrorHandler` | Público | Manejador centralizado de errores. Registra en Log y muestra mensaje al usuario. |
| `ProtegerTodoInicialmente` | Privado | Oculta todas las hojas (`xlSheetVeryHidden`) y oculta la cinta de Excel antes del login. |
| `RotarLogs` | Privado | Conserva solo los N registros más recientes para evitar crecimiento ilimitado del log. |
| `Workbook_BeforeClose` | Evento | Registra el cierre del libro en el log. |
| `Workbook_SheetActivate` | Evento | Registra cada cambio de hoja activa. |

---

## Modulo4_Sesion.vba
**Tipo:** Módulo estándar  
**Variables globales:** `UsuarioActual`, `NivelAcceso`

Sistema completo de autenticación y control de acceso por roles.

### Funciones/Procedimientos clave

| Nombre | Descripción |
|---|---|
| `IniciarSesion` | Muestra el cuadro de diálogo de usuario/contraseña. Controla intentos fallidos y determina el nivel de acceso. |
| `ReprotegerHojas` | Reaplica la protección a todas las hojas después de que las macros realizan cambios. |
| `GuardarAutomaticamente` | Programa el guardado del libro cada 3 minutos usando `Application.OnTime`. |
| `DetenerGuardadoAutomatico` | Cancela el guardado automático programado. |

### Flujo de autenticación
```
Workbook_Open
    └── ProtegerTodoInicialmente (oculta todas las hojas)
        └── IniciarSesion
            ├── Usuario 00 → Acceso total, todas las hojas visibles
            ├── Usuario 01 → Dashboard + Factura + Buscar
            └── Usuario 02 → Dashboard + Factura (solo combustible)
```

---

## Modulo5_Dashboard.vba
**Tipo:** Módulo estándar

Genera y actualiza todos los elementos visuales del panel de control.

### Funciones/Procedimientos clave

| Nombre | Descripción |
|---|---|
| `GenerarDashboardCompleto` | Macro principal que orquesta la creación completa del dashboard. |
| `CrearPanelDeControl` | Configura el área de controles (botones, validaciones, tasa Bs/USD). |
| `ActualizarDatos` | Actualiza los rangos de datos y refresca todos los gráficos. |
| `CrearTablasDinamicas` | Genera las tablas dinámicas que alimentan los gráficos. |
| `CrearSlicerTemporal` | Crea el segmentador de períodos (Hoy, Semanal, Mensual...). |
| `CrearSlicerTipoFactura` | Crea el segmentador por tipo de factura (Honorario, Combustible, H&C). |
| `CrearGraficosDashboard` | Genera los gráficos de barras (vertical y horizontal) y el gráfico circular. |
| `ActualizarCampoCalculado` | Recalcula el campo de equivalencia Bs/USD en las tablas dinámicas. |

---

## Hoja1_Dashboard.vba
**Tipo:** Módulo de hoja (`Dashboard`)

Maneja los eventos interactivos del panel de control.

### Funciones/Procedimientos clave

| Nombre | Descripción |
|---|---|
| `Worksheet_Change` | Detecta cambios en las celdas de configuración (E3, E4, E5 para tipo de gráfico; B6 para tasa) y actualiza la vista. |
| `MostrarOcultarGraficos` | Muestra u oculta gráficos específicos según la selección del panel de control. |
| `IsValidChartSelection` | Valida que el valor seleccionado corresponda a una opción de gráfico reconocida. |

---

## Hoja2_Factura.vba
**Tipo:** Módulo de hoja (`Factura`)

Toda la lógica del formulario de registro de facturas.

### Funciones/Procedimientos clave

| Nombre | Descripción |
|---|---|
| `Worksheet_Activate` | Al activar la hoja, inicializa todos los controles. |
| `InicializarControles` | Carga los ComboBoxes con datos desde las hojas `Datos` y `Extras`. Configura restricciones según nivel de usuario. |
| `GuardarFactura` | Valida y guarda un registro nuevo en la hoja `Facturas`. Incluye validaciones de campos obligatorios, formato de cédula y coherencia de datos de pago. |
| `cbxFactura_Change` | Evento del ComboBox de tipo de factura. Muestra u oculta campos según el tipo seleccionado (Honorario/Combustible/H&C). |
| `cbxMPago_Change` | Muestra u oculta los campos de datos bancarios según el método de pago seleccionado (Efectivo/Pagomóvil). |
| `CargarAutocompletado` | Al seleccionar un alumno ya registrado, autocompleta su cédula y datos vinculados. |

---

## Hoja3_Buscar.vba
**Tipo:** Módulo de hoja (`Buscar`)

Motor de búsqueda, edición y exportación de facturas.

### Funciones/Procedimientos clave

| Nombre | Descripción |
|---|---|
| `Worksheet_Activate` | Carga encabezados y opciones en los filtros al activar la hoja. |
| `BuscarFacturas` | Filtra la base de datos según los criterios seleccionados (campo + valor + período) y muestra los resultados en la tabla. |
| `GuardarCambios` | Guarda las modificaciones hechas directamente sobre los resultados mostrados en la tabla de búsqueda. |
| `GenerarPDF` | Exporta los registros visibles actualmente a un archivo PDF, con nombre automático basado en el filtro aplicado y la fecha. |
| `SincronizarBaseDeDatos` | Lee todos los archivos `.xlsm` en la misma carpeta y combina sus bases de datos. |
| `RestaurarBaseDeDatos` | Revierte la base de datos local desde el respaldo interno. |

---

## Modulo1_Validaciones.vba
**Tipo:** Módulo estándar  
**Propósito:** Funciones puras de validación de datos.

| Función | Descripción |
|---|---|
| `EsValidoNombreApellido` | Valida que el string tenga exactamente dos palabras (nombre y apellido). |
| `EsCedulaValida` | Valida formato de cédula venezolana (`V` o `E` seguido de números). |
| `EsMontoValido` | Verifica que el tipo de moneda sea `DIVISAS` o `BOLIVARES`. |
| `EsMetodoPagoValido` | Verifica que el método de pago sea `EFECTIVO` o `PAGOMOVIL`. |
| `CargarEnComboBox` | Utilitario genérico para poblar un ComboBox desde un rango de una hoja. |
| `MostrarCamposPagoMovil` | Muestra u oculta los OLEObjects de datos bancarios. |
| `QuickSort` | Implementación de QuickSort para ordenar arrays de strings. |

---

## Modulo2_Formularios.vba
**Tipo:** Módulo estándar  
**Propósito:** Operaciones sobre los controles ActiveX del formulario.

| Procedimiento | Descripción |
|---|---|
| `LimpiarCampos` | Resetea todos los ComboBoxes del formulario de factura a su estado vacío. |
| `LimpiarTablaDatosFactura` | Elimina todos los registros de la hoja `Facturas` (mantiene encabezado). |
| `LimpiarTablaDatosUsuarios` | Elimina los registros de usuarios desde la fila 3 en adelante. |
| `VerControlesHojaActiva` | Utilitario de depuración: lista en consola todos los OLEObjects de la hoja activa. |

---

## Modulo3_DatosPrueba.vba
**Tipo:** Módulo estándar  
**Propósito:** Generación de datos ficticios para pruebas internas.

| Procedimiento | Descripción |
|---|---|
| `InsertarDatosDePrueba` | Genera N facturas aleatorias con datos válidos pero ficticios. Solicita la cantidad mediante `InputBox`. Usa arrays de nombres, apellidos, aeronaves y bancos predefinidos. |
| `GenerarCedula` | Genera una cédula venezolana ficticia aleatoria (`V` + 8 dígitos). |
| `GenerarIDUnico` | Genera un identificador único para cada registro. |

---

## Notas de Implementación

### Protección de hojas con UserInterfaceOnly
Todas las hojas usan `Protect Password:="seguro", UserInterfaceOnly:=True`. Esto permite que las macros lean y escriban sin desproteger, pero el usuario no puede modificar celdas protegidas desde la interfaz.

### Ocultado de la cinta de opciones
```vba
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
```
Se usa el método `ExecuteExcel4Macro` por compatibilidad con versiones 2016+, ya que en Excel moderno no existe una propiedad directa para ocultar el Ribbon de forma no persistente.

### Manejo de eventos en cascada
Múltiples procedimientos desactivan `Application.EnableEvents = False` al inicio para evitar que los cambios programáticos disparen eventos secundarios no deseados (especialmente en la carga de ComboBoxes y la escritura en la base de datos).
