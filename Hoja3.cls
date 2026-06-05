'================== ESPACIO DE TRABAJO G-FELX ====================
Private Sub Worksheet_Activate()
    CargarDimensiones
    LlenarComboBox
    CargarColumna
    With Me.oleobjects("Orientaccion_BRR").Object
        .Clear
        .additem "Horizontal"
    End With
    With Me.oleobjects("Tipo_BRR").Object
        .Clear
        .additem "Tierra"
        .additem "Neutro"
        .additem "BusBar"
    End With
end sub
Private sub Tipo_BRR_Change()
    Dim tipo As String
    tipo = Me.oleobjects("Tipo_BRR").Object.Value
    with Me.oleobjects("Amperaje_BRR").Object
        .Clear
        Select case tipo
            case "Tierra", "Neutro"
                .additem "Defaut"
            case "BusBar"
                .additem "320" 
                .additem "400" 
        end select
    end with  
    with Me.oleobjects("Ancho_BRR").Object
        .Clear
        Select case tipo
            case "Tierra", "Neutro"
                .additem "720"
                .additem "400"
            case "BusBar"
                .additem "720"
        end select
    end with
End sub
public sub limpiar_area()
    ThisWorkbook.Sheets("frontal GFX").Range("JI96:MW187").Clear 
    ThisWorkbook.Sheets("frontal GFX").Range("JI96:MW187").Interior.COLOR = &HF2F2F2
end sub

Sub MoverBloque(ref As String, BD As String)
    Dim origen As Range
    Dim destino As Range
    Dim wsDatos As Worksheet
    Dim wsActual As Worksheet

    Set wsDatos = ThisWorkbook.Sheets(BD)
    Set wsActual = ActiveSheet

    On Error GoTo errHandler
    Set origen = wsDatos.Range(ref)
    Set destino = wsActual.Range("JK97")

    origen.Copy Destination:=destino

    'wsActual.Range(destino, destino.Offset(origen.Rows.Count - 1, origen.Columns.Count - 1)).Select
    Exit Sub

    errHandler:
        MsgBox "Referencia no encontrada ", vbCritical
End Sub

Public sub Buscar_envo1()
    ThisWorkbook.Sheets("frontal GFX").Range("JI96:MW187").Clear
    ThisWorkbook.Sheets("frontal GFX").Range("JI96:MW187").Interior.COLOR = &HF2F2F2
    dim ref As String
    ref = Me.oleobjects("DimensionesENV").Object.Value
    resultado = "FM_" & ref
    'MSgBox resultado
    MoverBloque resultado, "Envolventes_GFLEX"
End sub
public sub Buscar_barras()
    ThisWorkbook.Sheets("frontal GFX").Range("JI96:MW187").Clear
    ThisWorkbook.Sheets("frontal GFX").Range("JI96:MW187").Interior.COLOR = &HF2F2F2
    dim orientacion As String, tipo As String, amperaje As String, ancho As String
    orientacion = Me.oleobjects("Orientaccion_BRR").Object.Value
    tipo = Me.oleobjects("Tipo_BRR").Object.Value
    amperaje = Me.oleobjects("Amperaje_BRR").Object.Value
    ancho = Me.oleobjects("Ancho_BRR").Object.Value
    Select Case orientacion
        Case "Horizontal"
            orientacion = "H"
    end select
    select case tipo
        case "Tierra"
            tipo = "T"
        case "Neutro"
            tipo = "N"
        case "BusBar"
            tipo = "BB"
    end select
    Select Case amperaje
        Case "Defaut"
            amperaje = "NA"
    end select
    resultado = "G_" & orientacion & tipo & "X" & ancho & "X" & amperaje
   ' MSgBox resultado
    MoverBloque resultado, "HBB_TN"
End sub
Private Sub CargarColumna()
    Dim wscanasta As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim dic As Object
    
    Set wscanasta = ThisWorkbook.Sheets("EPanel_Master")
    Set dic = CreateObject("Scripting.Dictionary")
    
    ' Ajusta el 4 si tus datos empiezan en otra fila
    ultimaFila = wscanasta.Cells(wscanasta.Rows.Count, "A").End(xlUp).Row
    
    Me.Columna.Clear
    Me.Cubiculo.Clear
    Me.Equipo_canasta.Clear
    Me.Orientacion_canasta.Clear
    
    
    For i = 4 To ultimaFila
        If Trim(wscanasta.Cells(i, 1).Value) <> "" Then
            If Not dic.exists(wscanasta.Cells(i, 1).Value) Then
                dic.Add wscanasta.Cells(i, 1).Value, Nothing
                Me.Columna.AddItem wscanasta.Cells(i, 1).Value
            End If
        End If
    Next i
End Sub
Private Sub Columna_Change()
    Dim wscanasta As Worksheet
    Dim ultimaFila As Long 
    Dim i As Long
    Dim dic As Object
     
    Set wscanasta = ThisWorkbook.Sheets("EPanel_Master")
    Set dic = CreateObject("Scripting.Dictionary")
    
    ultimaFila = wscanasta.Cells(wscanasta.Rows.Count, "A").End(xlUp).Row
    
    Me.Cubiculo.Clear
    Me.Equipo_canasta.Clear
    
    For i = 4 To ultimaFila
        If wscanasta.Cells(i, 1).Value = Me.Columna.Value Then
            If Not dic.exists(wscanasta.Cells(i, 2).Value) Then
                dic.Add wscanasta.Cells(i, 2).Value, Nothing
                Me.Cubiculo.AddItem wscanasta.Cells(i, 2).Value
            End If
        End If
    Next i
End Sub
Private Sub Cubiculo_Change()
    Dim wscanasta As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim dic As Object
    
    Set wscanasta = ThisWorkbook.Sheets("EPanel_Master")
    Set dic = CreateObject("Scripting.Dictionary")
    
    ultimaFila = wscanasta.Cells(wscanasta.Rows.Count, "A").End(xlUp).Row
    
    Me.Equipo_canasta.Clear
    
    For i = 4 To ultimaFila
        If wscanasta.Cells(i, 1).Value = Me.Columna.Value And _
           wscanasta.Cells(i, 2).Value = Me.Cubiculo.Value Then
            If Not dic.exists(wscanasta.Cells(i, 3).Value) Then
                dic.Add wscanasta.Cells(i, 3).Value, Nothing
                Me.Equipo_canasta.AddItem wscanasta.Cells(i, 3).Value
            End If
        End If
    Next i
End Sub
Private Sub Equipo_canasta_Change()
    Dim wscanasta As Worksheet
    dim ultimaFila As Long
    Dim i As Long
    dim dic As Object
    
    Set wscanasta = ThisWorkbook.Sheets("EPanel_Master")
    set dic = CreateObject("Scripting.Dictionary")
    
    ultimaFila = wscanasta.Cells(wscanasta.Rows.Count, "A").End(xlUp).Row
    
    Me.Orientacion_canasta.Clear
    
    For i = 4 To ultimaFila
        If wscanasta.Cells(i, 1).Value = Me.Columna.Value And _
           wscanasta.Cells(i, 2).Value = Me.Cubiculo.Value And _
           wscanasta.Cells(i, 3).Value = Me.Equipo_canasta.Value Then
            If Not dic.exists(wscanasta.Cells(i, 4).Value) Then
                dic.Add wscanasta.Cells(i, 4).Value, Nothing
                Me.Orientacion_canasta.AddItem wscanasta.Cells(i, 4).Value
            End If
        End If
    Next i
End Sub
private sub Orientacion_canasta_Change()
    Dim dato1 as string
    dim dato2 as string
    dato1= Me.oleobjects("Equipo_canasta").Object.Value
    dato2= Me.oleobjects("Orientacion_canasta").Object.Value
    with Me.oleobjects("ancho_main").Object
        .clear
        select case dato1
            case "CVS250"
                Select Case dato2
                    case "Horizontal"
                        .additem "720"
                end select
        end select                
    end with
    with Me.oleobjects("sentido_main").Object
        .clear
        select case dato1
            case "CVS250"
                Select Case dato2
                    case "Horizontal"
                        .additem "Izquierda"
                end select
        end select
    end with
End sub

public sub Buscar_alamacen_main()
    thisworkbook.sheets("frontal GFX").range("JI96:MW187").clear
    thisworkbook.sheets("frontal GFX").range("JI96:MW187").interior.color = &hf2f2f2
    dim orientacion as string
    dim equipo as string
    dim ancho as string
    dim sentido as string
    orientacion = Me.oleobjects("Orientacion_canasta").Object.Value
    equipo = Me.oleobjects("Equipo_canasta").Object.Value
    ancho = Me.oleobjects("ancho_main").Object.Value
    sentido = Me.oleobjects("sentido_main").Object.Value
    select case orientacion
        case "Horizontal"
            orientacion = "HBB"
    end select
    select case sentido
        case "Izquierda"
            sentido = "IZQ"
    end select
    resultado = "G_" & orientacion & "XMainX" & equipo & "X" & ancho & "X" & sentido
    msgbox resultado
    moverbloque resultado, "TRAZADO_BARRAS"
End sub
