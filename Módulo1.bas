Attribute VB_Name = "M�dulo1"
'============== VARIBLES GLOBALES / RECOLECCIONJ DE DATOS ========================
Public DimensionesGFLX As String
Public resultado As String
'=================================================================================
Public Sub CargarDimensiones()
    DimensionesGFLX = Worksheets("PORTADA G-FLEX").Range("CP44").Value  ' Cambia A1 por tu celda
End Sub
Public Sub LlenarComboBox()
    Dim opciones() As String
    Dim i As Integer
    Dim ws As Worksheet
    Dim cbo As Object
    
    Set ws = ThisWorkbook.Worksheets("frontal GFX")
    Set cbo = ws.OLEObjects("DimensionesENV").Object  ' .Object accede al control ActiveX real
    
    cbo.Clear
    
    opciones = Split(DimensionesGFLX, ",")
    
    For i = 0 To UBound(opciones)
        cbo.AddItem Trim(opciones(i))
    Next i
End Sub
