Attribute VB_Name = "Módulo4"
'============ IMPORTAR LAYOUT DE T-POWER ============
' PREGUNTAR AL USUARIO LA UBICACION DEL ARCHIVO
' SELECCIONAR LA HOJA DE TRABAJO E IMPORTAR
' SOLO IMPORTAR EL RANGO E16:GT126 -> PEGAR EN BB42

Sub IMPORTARHOJAS()
    Dim rutaArchivo As Variant
    Dim hojaDestino As Worksheet
    Dim nombreHoja As String
    Dim listaHojas As String
    Dim hoja As Worksheet
    Dim LibroImportado As Workbook
    Dim HojaImportada As Worksheet

    ' Seleccionar archivo
    rutaArchivo = Application.GetOpenFilename("Archivos de Excel (*.xls*), *.xls*", , "Selecciona el archivo que deseas importar")
    If rutaArchivo = False Then Exit Sub

    ' Abrir archivo y mostrar hojas
    Set LibroImportado = Workbooks.Open(rutaArchivo)
    listaHojas = ""
    For Each hoja In LibroImportado.Sheets
        listaHojas = listaHojas & hoja.Name & vbCrLf
    Next hoja

    nombreHoja = InputBox("Selecciona una hoja:" & vbCrLf & listaHojas, "Seleccionar hoja")
    If nombreHoja = "" Then
        LibroImportado.Close False
        Exit Sub
    End If

    On Error Resume Next
    Set HojaImportada = LibroImportado.Sheets(nombreHoja)
    On Error GoTo 0

    If HojaImportada Is Nothing Then
        MsgBox "La hoja no existe.", vbCritical
        LibroImportado.Close False
        Exit Sub
    End If

    ' Copiar el rango fijo E16:GT126
    HojaImportada.Range("E16:GT126").Copy

    ' Pegar en hoja activa del archivo base, empezando en BB42
    Set hojaDestino = ThisWorkbook.ActiveSheet
    hojaDestino.Range("BB43").PasteSpecial xlPasteAll
    Application.CutCopyMode = False

    ' Cerrar el libro importado
    LibroImportado.Close False

    MsgBox "Layout extraido corectamente", vbInformation
End Sub