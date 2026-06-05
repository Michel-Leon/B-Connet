Attribute VB_Name = "Módulo6"
'==========================================================================
'  MACRO: UnirColores (version 2 celdas de ancho)
'  ----------------------------------------------------------------------
'  Detecta pares de BLOQUES de 2 celdas horizontales del mismo color
'  dentro del rango seleccionado y pinta el camino mas corto entre
'  cada par usando BFS, manteniendo siempre 2 celdas de ancho.
'
'  Esto simula la conexion de un interruptor (2 celdas) a un barraje
'  de cobre horizontal (2 celdas), tal como se ve en un tablero electrico.
'
'  USO:
'    1) Selecciona el rango donde estan las celdas de colores.
'    2) Ejecuta esta macro (Alt+F8 -> UnirColores -> Ejecutar).
'
'  CONFIGURACION:
'    Modifica el bloque "CONFIGURACION DE COLORES" mas abajo.
'==========================================================================

Public Sub UnirColores()
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox( _
        Prompt:="Selecciona el rango donde estan las celdas de colores:" & vbCrLf & _
                "(Los puntos de conexion deben ser bloques de 2 celdas horizontales)", _
        Title:="Unir colores - camino de 2 celdas de ancho", _
        Default:=Selection.Address, _
        Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub

    '======================================================================
    '  CONFIGURACION DE COLORES
    '  ---------------------------------------------------------------------
    '  Cada color que aparezca en EXACTAMENTE 2 bloques de 2 celdas
    '  horizontales sera conectado por el camino mas corto (2 celdas ancho).
    '======================================================================
    Dim colores() As Long
    ReDim colores(1 To 2)
    colores(1) = HexToColor("#C65911")   ' Naranja del ejemplo
    colores(2) = HexToColor("#BF8F00")   ' Dorado del ejemplo
    ' --- Agrega mas colores aqui si quieres:
    ' ReDim Preserve colores(1 To 3)
    ' colores(3) = RGB(0, 112, 192)      ' Azul, por ejemplo
    '======================================================================

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet
    Set ws = rng.Worksheet

    Dim r1 As Long, c1 As Long, nR As Long, nC As Long
    r1 = rng.Row
    c1 = rng.Column
    nR = rng.Rows.Count
    nC = rng.Columns.Count

    Dim i As Long
    Dim totalPares As Long, totalCeldasPintadas As Long
    totalPares = 0
    totalCeldasPintadas = 0

    For i = LBound(colores) To UBound(colores)
        Dim colorObj As Long
        colorObj = colores(i)

        '--------------------------------------------------------------
        '  Buscar bloques de 2 celdas horizontales con este color.
        '  Un "bloque" existe en (fila, colIzq) si AMBAS celdas
        '  (fila, colIzq) y (fila, colIzq+1) tienen el color.
        '  Se usa una marca para no contar la misma celda dos veces.
        '--------------------------------------------------------------
        Dim bloques() As Long   ' (1,*)=fila, (2,*)=columna izquierda
        Dim nBloques As Long
        ReDim bloques(1 To 2, 1 To nR * nC)
        nBloques = 0

        Dim marcado() As Boolean
        ReDim marcado(1 To nR, 1 To nC)

        Dim rr As Long, cc As Long
        For rr = 1 To nR
            For cc = 1 To nC - 1
                If Not marcado(rr, cc) Then
                    If ws.Cells(r1 + rr - 1, c1 + cc - 1).Interior.Color = colorObj And _
                       ws.Cells(r1 + rr - 1, c1 + cc).Interior.Color = colorObj Then
                        nBloques = nBloques + 1
                        bloques(1, nBloques) = rr
                        bloques(2, nBloques) = cc
                        marcado(rr, cc) = True
                        marcado(rr, cc + 1) = True
                    End If
                End If
            Next cc
        Next rr

        If nBloques = 2 Then
            '--------------------------------------------------------------
            '  BFS de bloque 1 a bloque 2 (cada nodo = posicion de bloque
            '  de 2 celdas: fila + columna izquierda del bloque).
            '--------------------------------------------------------------
            Dim camino As Variant
            camino = BFSCaminoDoble( _
                bloques(1, 1), bloques(2, 1), _
                bloques(1, 2), bloques(2, 2), _
                nR, nC)

            If Not IsEmpty(camino) Then
                Dim k As Long
                For k = LBound(camino, 2) To UBound(camino, 2)
                    Dim pR As Long, pC As Long
                    pR = camino(1, k)
                    pC = camino(2, k)
                    ' Pintar las 2 celdas del bloque
                    ws.Cells(r1 + pR - 1, c1 + pC - 1).Interior.Color = colorObj
                    ws.Cells(r1 + pR - 1, c1 + pC).Interior.Color = colorObj
                    totalCeldasPintadas = totalCeldasPintadas + 2
                Next k
                totalPares = totalPares + 1
            Else
                MsgBox "No se encontro camino para el color " & ColorToHex(colorObj) & _
                       "." & vbCrLf & "Verifique que haya espacio libre.", vbExclamation
            End If

        ElseIf nBloques > 2 Then
            MsgBox "El color " & ColorToHex(colorObj) & " aparece en " & nBloques & _
                   " bloques de 2 celdas." & vbCrLf & _
                   "Solo se unen pares (exactamente 2 bloques).", vbExclamation
        ElseIf nBloques = 1 Then
            MsgBox "El color " & ColorToHex(colorObj) & " solo tiene 1 bloque de 2 celdas." & vbCrLf & _
                   "Se necesitan exactamente 2 bloques para conectar.", vbExclamation
        ' nBloques = 0 -> el color no esta en el rango, simplemente se omite
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Listo." & vbCrLf & _
           "Pares unidos: " & totalPares & vbCrLf & _
           "Celdas pintadas (incluyendo extremos): " & totalCeldasPintadas, vbInformation
End Sub


'--------------------------------------------------------------------------
'  BFS para bloques de 2 celdas de ancho.
'
'  Cada NODO es (fila, colIzq) y representa las celdas:
'     (fila, colIzq)  y  (fila, colIzq + 1)
'
'  Movimientos permitidos (4 direcciones):
'     Arriba:   (fila - 1, colIzq)    -> ambas celdas suben 1 fila
'     Abajo:    (fila + 1, colIzq)    -> ambas celdas bajan 1 fila
'     Izquierda:(fila, colIzq - 1)    -> el bloque se desplaza 1 col a la izq
'     Derecha:  (fila, colIzq + 1)    -> el bloque se desplaza 1 col a la der
'
'  Limites:
'     fila entre 1 y nR
'     colIzq entre 1 y nC-1  (para que colIzq+1 <= nC)
'
'  Devuelve una matriz (2 x N) con (fila, colIzq) de cada paso del camino,
'  incluyendo origen y destino. Devuelve Empty si no hay camino.
'--------------------------------------------------------------------------
Private Function BFSCaminoDoble(ByVal sR As Long, ByVal sC As Long, _
                                 ByVal tR As Long, ByVal tC As Long, _
                                 ByVal nR As Long, ByVal nC As Long) As Variant

    ' Numero maximo de columnas izquierdas posibles: nC - 1
    Dim maxC As Long
    maxC = nC - 1
    If maxC < 1 Then
        BFSCaminoDoble = Empty
        Exit Function
    End If

    Dim visitado() As Boolean
    Dim padreR() As Long
    Dim padreC() As Long
    ReDim visitado(1 To nR, 1 To maxC)
    ReDim padreR(1 To nR, 1 To maxC)
    ReDim padreC(1 To nR, 1 To maxC)

    ' Inicializar padres a 0 (sin padre)
    Dim ir As Long, ic As Long
    For ir = 1 To nR
        For ic = 1 To maxC
            padreR(ir, ic) = 0
            padreC(ir, ic) = 0
        Next ic
    Next ir

    ' Cola BFS
    Dim colaR() As Long, colaC() As Long
    ReDim colaR(1 To nR * maxC)
    ReDim colaC(1 To nR * maxC)
    Dim head As Long, tail As Long
    head = 1: tail = 1

    colaR(tail) = sR
    colaC(tail) = sC
    visitado(sR, sC) = True
    padreR(sR, sC) = -1   ' marcador de origen
    padreC(sR, sC) = -1

    ' Direcciones: arriba, abajo, izquierda, derecha
    Dim dr(1 To 4) As Long, dc(1 To 4) As Long
    dr(1) = -1: dc(1) = 0    ' arriba
    dr(2) = 1:  dc(2) = 0    ' abajo
    dr(3) = 0:  dc(3) = -1   ' izquierda
    dr(4) = 0:  dc(4) = 1    ' derecha

    Dim encontrado As Boolean
    encontrado = (sR = tR And sC = tC)

    Do While head <= tail And Not encontrado
        Dim cr As Long, cc As Long
        cr = colaR(head): cc = colaC(head)
        head = head + 1

        Dim d As Long
        For d = 1 To 4
            Dim nr2 As Long, nc2 As Long
            nr2 = cr + dr(d)
            nc2 = cc + dc(d)

            ' Verificar limites: fila en [1..nR], colIzq en [1..maxC]
            If nr2 >= 1 And nr2 <= nR And nc2 >= 1 And nc2 <= maxC Then
                If Not visitado(nr2, nc2) Then
                    visitado(nr2, nc2) = True
                    padreR(nr2, nc2) = cr
                    padreC(nr2, nc2) = cc
                    tail = tail + 1
                    colaR(tail) = nr2
                    colaC(tail) = nc2
                    If nr2 = tR And nc2 = tC Then
                        encontrado = True
                        Exit For
                    End If
                End If
            End If
        Next d
    Loop

    If Not encontrado Then
        BFSCaminoDoble = Empty
        Exit Function
    End If

    ' Reconstruir camino desde destino hasta origen
    Dim caminoTmpR() As Long, caminoTmpC() As Long
    ReDim caminoTmpR(1 To nR * maxC)
    ReDim caminoTmpC(1 To nR * maxC)
    Dim n As Long
    n = 0
    Dim r As Long, c As Long
    r = tR: c = tC

    Do
        n = n + 1
        caminoTmpR(n) = r
        caminoTmpC(n) = c
        If padreR(r, c) = -1 And padreC(r, c) = -1 Then Exit Do
        Dim tmpR As Long, tmpC As Long
        tmpR = padreR(r, c)
        tmpC = padreC(r, c)
        r = tmpR
        c = tmpC
    Loop

    ' Invertir (de origen a destino)
    Dim resultado() As Long
    ReDim resultado(1 To 2, 1 To n)
    Dim j As Long
    For j = 1 To n
        resultado(1, j) = caminoTmpR(n - j + 1)
        resultado(2, j) = caminoTmpC(n - j + 1)
    Next j

    BFSCaminoDoble = resultado
End Function


'--------------------------------------------------------------------------
'  HexToColor: convierte "#RRGGBB" o "RRGGBB" al Long que Excel usa en
'  Interior.Color (que es &HBBGGRR&).
'--------------------------------------------------------------------------
Public Function HexToColor(ByVal hex As String) As Long
    Dim s As String
    s = Replace(hex, "#", "")
    s = Trim(s)
    If Len(s) <> 6 Then
        Err.Raise vbObjectError + 1000, "HexToColor", _
            "Formato invalido: '" & hex & "'. Usa '#RRGGBB' o 'RRGGBB'."
    End If
    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Mid(s, 1, 2))
    g = CLng("&H" & Mid(s, 3, 2))
    b = CLng("&H" & Mid(s, 5, 2))
    HexToColor = RGB(r, g, b)
End Function


'--------------------------------------------------------------------------
'  ColorToHex: convierte el Long de Interior.Color a "#RRGGBB" para mostrar
'--------------------------------------------------------------------------
Public Function ColorToHex(ByVal c As Long) As String
    Dim r As Long, g As Long, b As Long
    r = c Mod 256
    g = (c \ 256) Mod 256
    b = (c \ 65536) Mod 256
    ColorToHex = "#" & Right("00" & hex(r), 2) & _
                       Right("00" & hex(g), 2) & _
                       Right("00" & hex(b), 2)
End Function
