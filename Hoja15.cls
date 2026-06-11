'================== ESPACIO DE TRABAJO G-FELX ====================
Private Sub Worksheet_Activate()
    LlenarDimensionesENV Me.OLEObjects("DimensionesENV").Object
    CargarColumna
    with Me.oleobjects("Orientaccion_BRR").Object
        .Clear
        .additem "Horizontal"
        .additem "Vertical"
    End With
    'msgbox altura_portada
End Sub
Private Sub Orientaccion_BRR_Change()
    Dim dato1 As String
    dato1 = Me.OLEObjects("Orientaccion_BRR").Object.Value

    With Me.OLEObjects("Tipobrr_").Object
        .Clear
        Select Case dato1
            Case "Horizontal"
                .AddItem "Tierra"
                .AddItem "Neutro"
                .AddItem "BusBar"
            Case "Vertical"
                .AddItem "BusBar"
        End Select
    End With

    If dato1 = "Horizontal" Then
        Me.shapes("TEXT1").TextFrame.Characters.Text = "Ancho"
    Else
        Me.shapes("TEXT1").TextFrame.Characters.Text = "Fraccion"
    End If
End Sub
private Sub Tipobrr__Change()
    Dim dato1 As String
    Dim dato2 As String
    dato1 = Me.OLEObjects("Orientaccion_BRR").Object.Value
    dato2 = Me.OLEObjects("Tipobrr_").Object.Value
    With Me.OLEObjects("Amperaje_BRR").Object
        .Clear
        select case dato1
            case "Horizontal"
                Select Case dato2
                    Case "Tierra", "Neutro"
                        .AddItem "Defaut"
                    Case "BusBar"
                        .AddItem "1250"
                        .AddItem "1600"
                        .AddItem "2000"
                        .AddItem "2500"
                        .AddItem "3200"
                        .AddItem "4000"
                    end select
            case "Vertical"
                Select Case dato2 
                    Case "BusBar"
                        .AddItem "630" 
                        .AddItem "1250"
                        .AddItem "1600"
                        .AddItem "2000"
                        .AddItem "2500"
                end select
        end select        
    End With
End Sub 
private Sub Amperaje_BRR_Change()
    Dim dato1 As String
    Dim dato2 As String  
    dim dato3 As String
    dato1 = Me.OLEObjects("Orientaccion_BRR").Object.Value
    dato2 = Me.OLEObjects("Tipobrr_").Object.Value
    dato3 = Me.OLEObjects("Amperaje_BRR").Object.Value
    with Me.OLEObjects("Ancho_BRR").Object
        .Clear 
        select case dato1
            case "Horizontal"
                select case dato2
                    case "Tierra", "Neutro", "BusBar"
                        select case dato3
                            case "1250"
                                .AddItem "400"
                                .AddItem "720"
                            case "1600", "2000", "2500"
                                .AddItem "720"
                            case "3200" ,"4000"  
                                .AddItem "720"
                                .AddItem "1120"
                            case "Defaut"   
                                .addItem "400"
                                .addItem "720"
                                .addItem "1120"
                        end select
                end select
            case "Vertical"
                select case dato3
                    case "630", "1250", "1600", "2000", "2500"
                        .AddItem "1of4"
                        .AddItem "1of2"
                        .AddItem "3of4"
                        IF altura_portada = "2000" or altura_portada = "2200" Then
                            .AddItem "1"
                        end if
                end select
        end select
    end with
End Sub  
 public sub limpiar_area()
    ThisWorkbook.Sheets("frontal TPW").Range("JI96:MW187").Clear 
    ThisWorkbook.Sheets("frontal TPW").Range("JI96:MW187").Interior.COLOR = &HF2F2F2
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

    wsActual.Range(destino, destino.Offset(origen.Rows.Count - 1, origen.Columns.Count - 1)).Select
    Exit Sub

    errHandler:
        MsgBox "Referencia no encontrada ", vbCritical
End Sub
Public sub Buscar_envo1()
    ThisWorkbook.Sheets("frontal TPW").Range("JI96:MW187").Clear
    ThisWorkbook.Sheets("frontal TPW").Range("JI96:MW187").Interior.COLOR = &HF2F2F2
    dim ref As String
    ref = Me.oleobjects("DimensionesENV").Object.Value
    resultado_tpower = "E" & ref
    MSgBox resultado_tpower
    MoverBloque resultado_tpower, "Envolventes_TPOWER"
End sub
public sub Buscar_Barraje()
    ThisWorkbook.Sheets("frontal TPW").Range("JI96:MW187").Clear
    ThisWorkbook.Sheets("frontal TPW").Range("JI96:MW187").Interior.COLOR = &HF2F2F2
    dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    dato1 = Me.oleobjects("Orientaccion_BRR").Object.Value
    dato2 = Me.oleobjects("Tipobrr_").Object.Value
    dato3 = Me.oleobjects("Amperaje_BRR").Object.Value
    dato4 = Me.oleobjects("Ancho_BRR").Object.Value
    select case dato1
        case "Horizontal":dato1 = "H"
        case "Vertical":dato1 = "V"
    end select 
    select case dato2
        case "Tierra":dato2 = "T"
        case "Neutro":dato2 = "N"
        case "BusBar":dato2 = "BB"
    end select
    If dato1 = "H" And (dato2 = "T" Or dato2 = "N") Then
        resultado_tpower = "T_" & dato1 & dato2 & "X" & dato4
        MoverBloque resultado_tpower, "HBB_TN"
    ElseIf dato1 = "H" And dato2 = "BB" Then
        resultado_tpower = "T_" & dato1 & dato2 & "X" & dato4 & "X" & dato3
        MoverBloque resultado_tpower, "HBB_TN"
    ElseIf dato1 = "V" And dato2 = "BB" Then
        resultado_tpower = "T_" & dato1 & dato2 & "X" & dato3 & "X" & dato4 & "X" & altura_portada
        MoverBloque resultado_tpower, "VBB"
    End If
    MSgBox resultado_tpower
   
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
            case "CVS250","EZC250","NSX250"
                Select Case dato2
                    case "Horizontal","Vertical","V. Extendida"
                        .additem "720"
                    Case else
                        .additem "No existe bloque en BD"    
                end select
            case "EZC100","GoPact100","GoPact250","CVS630","EZC630","GoPact630","NSX630","NS"
                Select Case dato2   
                    case "Horizontal"
                        .additem "720"
                     Case else
                        .additem "No existe bloque en BD"    
                End Select
            case "MTZ Fijo" ,"MTZ Extraible","NW Fijo","NW Extraible"
                Select Case dato2   
                    case "Vertical"
                        .additem "720"
                     Case else
                        .additem "No existe bloque en BD"    
                End Select
        end select                
    end with
End sub
private sub ancho_main_Change()
    Dim dato1 as string
    dim dato2 as string
    dim dato3 as string
    dato1= Me.oleobjects("Equipo_canasta").Object.Value
    dato2= Me.oleobjects("Orientacion_canasta").Object.Value
    dato3= Me.oleobjects("ancho_main").Object.Value
    with Me.oleobjects("sentido_main").Object
        .clear
        select case dato1
            case "CVS250","EZC250","NSX250","NS"
                Select Case dato2
                    case "Horizontal"
                        select case dato3
                            case "720" 
                                .additem "Izquierda"
                                .additem "Derecha"
                        end select
                    case "Vertical", "V. Extendida"
                        select case dato3
                            case "720"
                                .additem "Superior"
                                .additem "Inferior"
                        end select 
                end select         
            case "EZC100","GoPact100","GoPact250","CVS630","EZC630","GoPact630","NSX630","NS"
                Select Case dato2
                    case "Horizontal"
                        select case dato3 
                            case "720"
                                .additem "Izquierda"
                                .additem "Derecha"
                        end select           
                end select
            case "MTZ Fijo" ,"MTZ Extraible","NW Fijo","NW Extraible"
                Select Case dato2   
                    case "Vertical"
                        select case dato3
                            case "720"
                                .additem "Superior"
                                .additem "Inferior"
                        end select
                End Select    
        end select
    end with
end sub
public sub Buscar_alamacen_main()
    thisworkbook.sheets("frontal TPW").range("JI96:MW187").clear
    ThisWorkbook.Sheets("frontal TPW").Range("JI96:MW187").Interior.COLOR = &HF2F2F2
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
        case "Vertical","V. Extendida"
            orientacion = "VBB"
    end select
    select case sentido
        case "Izquierda"
            sentido = "IZQ"
        case "Derecha"
            sentido = "DRC"
        case "Superior"
            sentido = "SPR"
        case "Inferior"
            sentido = "INF"
    end select
    select case equipo
        case "NW Fijo":equipo = "NWFijo"
        case "NW Extraible":equipo = "NWExtraible"
        case "MTZ Fijo":equipo = "MTZFijo"
        case "MTZ Extraible":equipo = "MTZExtraible"
    end select
    resultado_tpower = "T_" & orientacion & "XMainX" & equipo & "X" & ancho & "X" & sentido
    msgbox resultado_tpower
    moverbloque resultado_tpower, "TRAZADO_BARRAS"
End sub