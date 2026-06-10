'================== ESPACIO DE TRABAJO G-FELX ====================
Private Sub Worksheet_Activate()
    LlenarDimensionesENV Me.OLEObjects("DimensionesENV").Object
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