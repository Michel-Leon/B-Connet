Attribute VB_Name = "Cargar_Maestros"

Option Explicit

'=========================================================
' MACRO: Cargar_Desde_Maestro
' Aplicacion: B-Connet
' Lee del MAESTRO_<ID>.xlsx las hojas _META, Cabecera y EPanel
' Permite: (1) buscar por ID directo  (2) elegir de lista completa
'=========================================================
Sub Cargar_Desde_Maestro()

    '----------------------------------------------------
    ' CONFIGURACION
    '----------------------------------------------------
    Const RUTA_SERVIDOR As String = "\\192.168.100.170\8.ape\APE_EXTERNA\0.M PROYECTOS\NUEVOSPROYECTOS\"
    Const PREFIJO_MAESTRO As String = "MAESTRO_"
    Const EXTENSION_MAESTRO As String = ".xlsx"
    Const SEPARADOR_ID As String = " - "

    '----------------------------------------------------
    ' VARIABLES
    '----------------------------------------------------
    Dim wbMaestro As Workbook
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim idProyecto As String
    Dim nombreCarpetaSeleccionada As String
    Dim rutaCarpeta As String
    Dim rutaArchivo As String
    Dim reporte As String
    Dim i As Long
    Dim hojaInfo As Variant
    Dim hojasImportar As Variant
    Dim lineaProducto As String

    '----------------------------------------------------
    ' PASO 1: Validar que el servidor sea accesible
    '----------------------------------------------------
    Application.StatusBar = "Verificando conexion al servidor..."

    If Dir(RUTA_SERVIDOR, vbDirectory) = "" Then
        Application.StatusBar = False
        MsgBox "No se puede acceder a la ruta del servidor:" & vbCrLf & vbCrLf & _
               RUTA_SERVIDOR & vbCrLf & vbCrLf & _
               "Verifique conexion y permisos.", vbCritical, "Error de conexion"
        Exit Sub
    End If

    '----------------------------------------------------
    ' PASO 2: Listar TODAS las carpetas con maestro valido
    '         Se hace en DOS PASADAS para no romper Dir()
    '----------------------------------------------------
    Application.StatusBar = "Buscando proyectos disponibles..."

    Dim carpetasTodas() As String
    Dim numTodas As Long
    ReDim carpetasTodas(1 To 2000)
    numTodas = 0

    Dim nombreCarpeta As String
    nombreCarpeta = Dir(RUTA_SERVIDOR & "*", vbDirectory)
    Do While nombreCarpeta <> ""
        If nombreCarpeta <> "." And nombreCarpeta <> ".." Then
            numTodas = numTodas + 1
            carpetasTodas(numTodas) = nombreCarpeta
        End If
        nombreCarpeta = Dir()
    Loop

    Dim carpetas() As String
    Dim ids() As String
    Dim numCarpetas As Long
    ReDim carpetas(1 To numTodas + 1)
    ReDim ids(1 To numTodas + 1)
    numCarpetas = 0

    Dim atributos As Long
    Dim posSeparador As Long
    Dim idExtraido As String
    Dim rutaMaestroEsperado As String
    Dim esDirectorio As Boolean

    For i = 1 To numTodas
        esDirectorio = False
        On Error Resume Next
        atributos = GetAttr(RUTA_SERVIDOR & carpetasTodas(i))
        If Err.Number = 0 Then
            If (atributos And vbDirectory) = vbDirectory Then esDirectorio = True
        End If
        Err.Clear
        On Error GoTo 0

        If esDirectorio Then
            posSeparador = InStr(1, carpetasTodas(i), SEPARADOR_ID)
            If posSeparador > 0 Then
                idExtraido = Trim(Left(carpetasTodas(i), posSeparador - 1))
                rutaMaestroEsperado = RUTA_SERVIDOR & carpetasTodas(i) & "\" & _
                                      PREFIJO_MAESTRO & idExtraido & EXTENSION_MAESTRO

                If Dir(rutaMaestroEsperado) <> "" Then
                    numCarpetas = numCarpetas + 1
                    carpetas(numCarpetas) = carpetasTodas(i)
                    ids(numCarpetas) = idExtraido
                End If
            End If
        End If
    Next i

    Application.StatusBar = False

    If numCarpetas = 0 Then
        MsgBox "No se encontraron proyectos con archivo maestro valido en:" & vbCrLf & _
               RUTA_SERVIDOR, vbExclamation, "Sin proyectos"
        Exit Sub
    End If

    '----------------------------------------------------
    ' PASO 3: Preguntar modo de busqueda
    '----------------------------------------------------
    Dim modoBusqueda As String
    Do
        modoBusqueda = Trim(InputBox( _
            "Se encontraron " & numCarpetas & " proyecto(s) disponibles." & vbCrLf & vbCrLf & _
            "Como desea seleccionar el proyecto?" & vbCrLf & vbCrLf & _
            "  1 = Buscar por ID" & vbCrLf & _
            "  2 = Ver lista completa", _
            "Modo de busqueda"))
        If modoBusqueda = "" Then Exit Sub
    Loop While modoBusqueda <> "1" And modoBusqueda <> "2"

    Dim indiceSeleccionado As Long
    indiceSeleccionado = 0

    '----------------------------------------------------
    ' PASO 3A: BUSQUEDA POR ID DIRECTO
    '----------------------------------------------------
    If modoBusqueda = "1" Then
        Dim idIngresado As String
        idIngresado = Trim(CStr(InputBox( _
            "Ingrese el ID del proyecto:" & vbCrLf & vbCrLf & _
            "El ID es el numero antes del guion en el nombre de la carpeta." & vbCrLf & _
            "Ejemplo: para '1245-5903 - CTS' el ID es '1245-5903'", _
            "Buscar por ID")))

        If idIngresado = "" Then Exit Sub

        For i = 1 To numCarpetas
            If UCase(ids(i)) = UCase(idIngresado) Then
                indiceSeleccionado = i
                Exit For
            End If
        Next i

        If indiceSeleccionado = 0 Then
            Dim coincidencias() As Long
            Dim numCoinc As Long
            ReDim coincidencias(1 To numCarpetas)
            numCoinc = 0

            For i = 1 To numCarpetas
                If InStr(1, ids(i), idIngresado, vbTextCompare) > 0 Or _
                   InStr(1, carpetas(i), idIngresado, vbTextCompare) > 0 Then
                    numCoinc = numCoinc + 1
                    coincidencias(numCoinc) = i
                End If
            Next i

            If numCoinc = 0 Then
                MsgBox "No se encontro ningun proyecto con el ID '" & idIngresado & "'." & vbCrLf & vbCrLf & _
                       "Verifique el ID e intente nuevamente.", vbExclamation, "Sin coincidencias"
                Exit Sub
            ElseIf numCoinc = 1 Then
                indiceSeleccionado = coincidencias(1)
            Else
                Dim listaCoinc As String
                listaCoinc = "Se encontraron " & numCoinc & " coincidencias para '" & _
                             idIngresado & "':" & vbCrLf & vbCrLf
                For i = 1 To numCoinc
                    listaCoinc = listaCoinc & "  " & i & " = " & carpetas(coincidencias(i)) & vbCrLf
                Next i
                listaCoinc = listaCoinc & vbCrLf & "Ingrese el numero (1 a " & numCoinc & "):"

                Dim selCoinc As String
                selCoinc = Trim(InputBox( _
                    Prompt:=listaCoinc, _
                    Title:="Seleccionar entre coincidencias"))

                If selCoinc = "" Then Exit Sub
                If Not IsNumeric(selCoinc) Then
                    MsgBox "Ingrese un numero valido.", vbExclamation
                    Exit Sub
                End If
                Dim numSelCoinc As Long
                numSelCoinc = CLng(selCoinc)
                If numSelCoinc < 1 Or numSelCoinc > numCoinc Then
                    MsgBox "Numero fuera de rango. Debe ser entre 1 y " & numCoinc & ".", vbExclamation
                    Exit Sub
                End If
                indiceSeleccionado = coincidencias(numSelCoinc)
            End If
        End If

    '----------------------------------------------------
    ' PASO 3B: SELECCION POR LISTA COMPLETA
    '----------------------------------------------------
    Else
        Call OrdenarParesAlfabeticamente(carpetas, ids, numCarpetas)

        Dim listaTexto As String
        listaTexto = "Proyectos disponibles (" & numCarpetas & "):" & vbCrLf & vbCrLf

        For i = 1 To numCarpetas
            listaTexto = listaTexto & "  " & i & " = " & carpetas(i) & vbCrLf
        Next i

        listaTexto = listaTexto & vbCrLf & "Ingrese el numero (1 a " & numCarpetas & "):"

        Dim seleccion As String
        seleccion = Trim(InputBox( _
            Prompt:=listaTexto, _
            Title:="Seleccionar proyecto maestro"))

        If seleccion = "" Then Exit Sub
        If Not IsNumeric(seleccion) Then
            MsgBox "Ingrese un numero valido.", vbExclamation
            Exit Sub
        End If
        Dim numSel As Long
        numSel = CLng(seleccion)
        If numSel < 1 Or numSel > numCarpetas Then
            MsgBox "Numero fuera de rango. Debe ser entre 1 y " & numCarpetas & ".", vbExclamation
            Exit Sub
        End If

        indiceSeleccionado = numSel
    End If

    '----------------------------------------------------
    ' PASO 4: Construir rutas finales
    '----------------------------------------------------
    nombreCarpetaSeleccionada = carpetas(indiceSeleccionado)
    idProyecto = ids(indiceSeleccionado)
    rutaCarpeta = RUTA_SERVIDOR & nombreCarpetaSeleccionada & "\"
    rutaArchivo = rutaCarpeta & PREFIJO_MAESTRO & idProyecto & EXTENSION_MAESTRO

    '----------------------------------------------------
    ' PASO 5: Seleccionar linea de producto (con validacion)
    '----------------------------------------------------
    Do
        lineaProducto = Trim(InputBox( _
            "Proyecto seleccionado:" & vbCrLf & _
            "  " & nombreCarpetaSeleccionada & vbCrLf & vbCrLf & _
            "Seleccione la linea de producto:" & vbCrLf & vbCrLf & _
            "  1 = T-Power" & vbCrLf & _
            "  2 = G-Flex", _
            "Linea de producto"))
        If lineaProducto = "" Then Exit Sub
    Loop While lineaProducto <> "1" And lineaProducto <> "2"

    If lineaProducto = "1" Then
        hojasImportar = Array( _
            Array("_META", "Meta_Master", "B4:B16"), _
            Array("Cabecera", "Cabecera_Master_TPower", "A3:Z100"), _
            Array("EPanel", "EPanel_Master", "A4:Z100") _
        )
    Else
        hojasImportar = Array( _
            Array("_META", "Meta_Master", "B4:B16"), _
            Array("Cabecera", "Cabecera_Master_GFlex", "A3:Z100"), _
            Array("EPanel", "EPanel_Master", "A4:Z100") _
        )
    End If

    '----------------------------------------------------
    ' PASO 6: Validar hojas locales y avisar si hay datos previos
    '----------------------------------------------------
    Dim tieneDatos As Boolean
    tieneDatos = False

    For Each hojaInfo In hojasImportar
        Set wsDestino = Nothing
        On Error Resume Next
        Set wsDestino = ThisWorkbook.Sheets(hojaInfo(1))
        On Error GoTo 0

        If wsDestino Is Nothing Then
            MsgBox "No se encontro la hoja local '" & hojaInfo(1) & "' en este archivo." & vbCrLf & vbCrLf & _
                   "Verifique que la hoja exista antes de continuar.", vbCritical, "Hoja no encontrada"
            Exit Sub
        End If

        If Application.WorksheetFunction.CountA(wsDestino.Range(hojaInfo(2))) > 0 Then
            tieneDatos = True
        End If
    Next hojaInfo

    If tieneDatos Then
        Dim respuesta As VbMsgBoxResult
        respuesta = MsgBox( _
            "Las hojas locales ya contienen datos." & vbCrLf & vbCrLf & _
            "Al continuar, los datos actuales seran reemplazados con los del proyecto:" & vbCrLf & vbCrLf & _
            "  " & nombreCarpetaSeleccionada & vbCrLf & vbCrLf & _
            "Desea continuar?", _
            vbYesNo + vbQuestion, "Confirmar reemplazo de datos")

        If respuesta = vbNo Then Exit Sub
    End If

    '----------------------------------------------------
    ' PASO 7: Abrir maestro en modo solo lectura
    '----------------------------------------------------
    On Error GoTo ManejoError

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Abriendo maestro del proyecto " & idProyecto & "..."

    Set wbMaestro = Workbooks.Open(Filename:=rutaArchivo, ReadOnly:=True, UpdateLinks:=0)

    Dim nombreLinea As String
    nombreLinea = IIf(lineaProducto = "1", "T-Power", "G-Flex")

    reporte = String(50, "-") & vbCrLf & _
              " PROYECTO : " & nombreCarpetaSeleccionada & vbCrLf & _
              " ID       : " & idProyecto & vbCrLf & _
              " LINEA    : " & nombreLinea & vbCrLf & _
              String(50, "-") & vbCrLf & vbCrLf & _
              " HOJAS IMPORTADAS:" & vbCrLf & vbCrLf

    '----------------------------------------------------
    ' PASO 8: Importar cada hoja con rango exacto
    '----------------------------------------------------
    For Each hojaInfo In hojasImportar
        Dim nombreHojaOrigen As String
        Dim nombreHojaDestino As String
        Dim rangoCopia As String

        nombreHojaOrigen = hojaInfo(0)
        nombreHojaDestino = hojaInfo(1)
        rangoCopia = hojaInfo(2)

        Application.StatusBar = "Importando hoja: " & nombreHojaOrigen & "..."

        Set wsOrigen = Nothing
        On Error Resume Next
        Set wsOrigen = wbMaestro.Sheets(nombreHojaOrigen)
        On Error GoTo ManejoError

        If wsOrigen Is Nothing Then
            reporte = reporte & " [---] " & nombreHojaOrigen & ": no encontrada en el maestro" & vbCrLf
            GoTo SiguienteHoja
        End If

        Set wsDestino = ThisWorkbook.Sheets(nombreHojaDestino)
        wsDestino.Range(rangoCopia).ClearContents
        wsDestino.Range(rangoCopia).Value = wsOrigen.Range(rangoCopia).Value

        Dim celdasConDato As Long
        celdasConDato = Application.WorksheetFunction.CountA(wsDestino.Range(rangoCopia))

        reporte = reporte & " [OK]  " & nombreHojaOrigen & " -> " & nombreHojaDestino & _
                  "  (" & celdasConDato & " celda(s))" & vbCrLf

SiguienteHoja:
    Next hojaInfo

    '----------------------------------------------------
    ' PASO 9: Cerrar maestro sin guardar
    '----------------------------------------------------
    wbMaestro.Close SaveChanges:=False
    Set wbMaestro = Nothing

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False

    On Error Resume Next
    ThisWorkbook.Sheets("Config").Range("ID_Proyecto_Cargado").Value = idProyecto
    On Error GoTo 0

    MsgBox reporte & vbCrLf & String(50, "-") & vbCrLf & _
           " Carga completada exitosamente.", vbInformation, "Importacion finalizada"
    Exit Sub

ManejoError:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False

    If Not wbMaestro Is Nothing Then
        wbMaestro.Close SaveChanges:=False
    End If

    MsgBox "Ocurrio un error durante la importacion." & vbCrLf & vbCrLf & _
           "Detalle : " & Err.Description & vbCrLf & _
           "Numero  : " & Err.Number, vbCritical, "Error de importacion"
End Sub


'=========================================================
' Helper: ordena alfabeticamente dos arrays paralelos
' (carpetas y sus IDs) usando bubble sort
'=========================================================
Private Sub OrdenarParesAlfabeticamente(ByRef carpetas() As String, _
                                        ByRef ids() As String, _
                                        ByVal n As Long)
    Dim i As Long, j As Long
    Dim tmp As String

    For i = 1 To n - 1
        For j = 1 To n - i
            If StrComp(carpetas(j), carpetas(j + 1), vbTextCompare) > 0 Then
                tmp = carpetas(j): carpetas(j) = carpetas(j + 1): carpetas(j + 1) = tmp
                tmp = ids(j): ids(j) = ids(j + 1): ids(j + 1) = tmp
            End If
        Next j
    Next i
End Sub
 