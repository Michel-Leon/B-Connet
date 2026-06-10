Attribute VB_Name = "M�dulo3"
'varibles globales T-POwer
public altura_portada as string
public resultado_tpower as string
Public Sub LlenarDimensionesENV(cbo As MSForms.ComboBox)
    Dim ws As Worksheet
    Dim altura As Double
    Dim anchoTotal As Double
    Dim anchos As Variant
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("PORTADA T-POWER")

    ' Conversión segura por si las celdas tienen texto
    altura = Val(ws.Range("CM41").Value)
    anchoTotal = Val(ws.Range("CM47").Value)
    altura_portada = altura
    anchos = Array(200, 400, 720,1120)

    cbo.Clear

    For i = LBound(anchos) To UBound(anchos)
        If anchos(i) <= anchoTotal Then
            cbo.AddItem altura & "X" & anchos(i)
        End If
    Next i
End Sub