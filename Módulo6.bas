Attribute VB_Name = "Módulo6"
'==========================================================================
'  MACRO: UnirColores (version 2x2)
'  El camino es un bloque de 2x2 celdas que se mueve de a 1 celda
'  en 4 direcciones. Pinta sobre cualquier color existente.
'  Los puntos de inicio/destino pueden ser 2 celdas horiz o vert.
'==========================================================================

Public Sub UnirColores()

    '----------------------------------------------------------------------
    '  PASO 1: Bloque de INICIO
    '----------------------------------------------------------------------
    Dim rngInicio As Range
    On Error Resume Next
    Set rngInicio = Application.InputBox( _
        Prompt:="PASO 1 de 3" & vbCrLf & vbCrLf & _
                "Selecciona el bloque de INICIO:" & vbCrLf & _
                "(2 celdas horizontales o verticales contiguas)", _
        Title:="Unir colores — Punto de INICIO", _
        Default:=Selection.Address, _
        Type:=8)
    On Error GoTo 0
    If rngInicio Is Nothing Then Exit Sub

    If Not EsBloque2Celdas(rngInicio) Then
        MsgBox "El bloque de INICIO debe ser exactamente 2 celdas contiguas " & _
               "(horizontales o verticales).", vbExclamation, "Seleccion invalida"
        Exit Sub
    End If

    Dim colorInicio As Long
    colorInicio = rngInicio.Cells(1, 1).Interior.Color

    '----------------------------------------------------------------------
    '  PASO 2: Bloque de DESTINO
    '----------------------------------------------------------------------
    Dim rngDestino As Range
    On Error Resume Next
    Set rngDestino = Application.InputBox( _
        Prompt:="PASO 2 de 3" & vbCrLf & vbCrLf & _
                "Selecciona el bloque de DESTINO:" & vbCrLf & _
                "(2 celdas horizontales o verticales contiguas)", _
        Title:="Unir colores — Punto de DESTINO", _
        Type:=8)
    On Error GoTo 0
    If rngDestino Is Nothing Then Exit Sub

    If Not EsBloque2Celdas(rngDestino) Then
        MsgBox "El bloque de DESTINO debe ser exactamente 2 celdas contiguas " & _
               "(horizontales o verticales).", vbExclamation, "Seleccion invalida"
        Exit Sub
    End If

    '----------------------------------------------------------------------
    '  PASO 3: Area de trabajo
    '  Se sugiere automaticamente el rectangulo minimo que contiene
    '  ambos bloques. El usuario puede ampliarlo.
    '----------------------------------------------------------------------
    Dim ws As Worksheet
    Set ws = rngInicio.Worksheet

    Dim aR1 As Long, aC1 As Long, aR2 As Long, aC2 As Long
    aR1 = Application.Min(rngInicio.Row, rngDestino.Row)
    aC1 = Application.Min(rngInicio.Column, rngDestino.Column)
    aR2 = Application.Max(rngInicio.Row + rngInicio.Rows.Count - 1, _
                        rngDestino.Row + rngDestino.Rows.Count - 1)
    aC2 = Application.Max(rngInicio.Column + rngInicio.Columns.Count - 1, _
                        rngDestino.Column + rngDestino.Columns.Count - 1)

    '  Garantizar minimo 2 filas y 2 columnas en el area sugerida
    '  y ademas agregar 1 fila/columna de margen en cada extremo
    '  para que el bloque 2x2 pueda posicionarse sobre los extremos
    aR1 = aR1 - 1
    aC1 = aC1 - 1
    aR2 = aR2 + 1
    aC2 = aC2 + 1

    Dim rngArea As Range
    On Error Resume Next
    Set rngArea = Application.InputBox( _
        Prompt:="PASO 3 de 3" & vbCrLf & vbCrLf & _
                "Define el AREA DE TRABAJO donde puede circular el camino." & vbCrLf & _
                "Puedes ampliarla si el camino necesita mas espacio.", _
        Title:="Unir colores — Area de trabajo", _
        Default:=ws.Range(ws.Cells(aR1, aC1), ws.Cells(aR2, aC2)).Address, _
        Type:=8)
    On Error GoTo 0
    If rngArea Is Nothing Then Exit Sub

    If Not EstaEnRango(rngInicio, rngArea) Then
        MsgBox "El bloque de INICIO esta fuera del area de trabajo." & vbCrLf & _
               "Amplia el area para que incluya ambos bloques.", vbExclamation
        Exit Sub
    End If
    If Not EstaEnRango(rngDestino, rngArea) Then
        MsgBox "El bloque de DESTINO esta fuera del area de trabajo." & vbCrLf & _
               "Amplia el area para que incluya ambos bloques.", vbExclamation
        Exit Sub
    End If

    '----------------------------------------------------------------------
    '  Coordenadas relativas al area de trabajo
    '  Cada nodo BFS = (filaTop, colLeft) de un bloque 2x2
    '  Limites validos: filaTop en [1..nR-1], colLeft en [1..nC-1]
    '----------------------------------------------------------------------
    Dim r1 As Long, c1 As Long, nR As Long, nC As Long
    r1 = rngArea.Row
    c1 = rngArea.Column
    nR = rngArea.Rows.Count
    nC = rngArea.Columns.Count

    If nR < 2 Or nC < 2 Then
        MsgBox "El area de trabajo debe tener al menos 2 filas y 2 columnas.", vbExclamation
        Exit Sub
    End If

    '  Posicion superior-izquierda de cada bloque en coords relativas
    Dim sR As Long, sC As Long, tR As Long, tC As Long
    ObtenerPosBloque rngInicio,  r1, c1, sR, sC
    ObtenerPosBloque rngDestino, r1, c1, tR, tC

    '  Para el BFS 2x2 el nodo es la esquina superior-izquierda del cuadrado.
    '  Si el bloque de inicio es vertical (col unica, 2 filas) la esquina
    '  superior-izq ya esta bien. Si es horizontal (fila unica, 2 cols)
    '  tambien. En ambos casos usamos (sR, sC) tal cual.
    '  Solo verificamos que queden dentro de [1..nR-1] x [1..nC-1].
    If sR < 1 Or sR > nR - 1 Or sC < 1 Or sC > nC - 1 Then
        MsgBox "El bloque de INICIO no cabe como 2x2 dentro del area." & vbCrLf & _
               "Amplia el area al menos 1 fila y 1 columna.", vbExclamation
        Exit Sub
    End If
    If tR < 1 Or tR > nR - 1 Or tC < 1 Or tC > nC - 1 Then
        MsgBox "El bloque de DESTINO no cabe como 2x2 dentro del area." & vbCrLf & _
               "Amplia el area al menos 1 fila y 1 columna.", vbExclamation
        Exit Sub
    End If

    '----------------------------------------------------------------------
    '  BFS y pintado
    '----------------------------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim camino As Variant
    camino = BFSCamino2x2(sR, sC, tR, tC, nR, nC)

    If Not IsEmpty(camino) Then
        Dim k As Long
        For k = LBound(camino, 2) To UBound(camino, 2)
            Dim pR As Long, pC As Long
            pR = camino(1, k)
            pC = camino(2, k)
            '  Pintar el bloque 2x2 con el color del inicio
            '  (sobreescribe cualquier color existente)
            ws.Cells(r1 + pR - 1,     c1 + pC - 1).Interior.Color = colorInicio
            ws.Cells(r1 + pR - 1,     c1 + pC    ).Interior.Color = colorInicio
            ws.Cells(r1 + pR,         c1 + pC - 1).Interior.Color = colorInicio
            ws.Cells(r1 + pR,         c1 + pC    ).Interior.Color = colorInicio
        Next k

        MsgBox "Conexion realizada." & vbCrLf & _
               "Pasos del camino: " & (UBound(camino, 2) - LBound(camino, 2) + 1) & vbCrLf & _
               "Celdas pintadas: "  & (UBound(camino, 2) - LBound(camino, 2) + 1) * 4, _
               vbInformation, "Listo"
    Else
        MsgBox "No se encontro un camino entre los dos bloques." & vbCrLf & _
               "Verifica que el area de trabajo sea suficientemente grande.", _
               vbExclamation, "Sin camino"
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


'--------------------------------------------------------------------------
'  BFS 2x2
'  Cada nodo = (filaTop, colLeft) = esquina superior-izquierda del 2x2.
'  Representa las 4 celdas:
'     (filaTop,   colLeft)   (filaTop,   colLeft+1)
'     (filaTop+1, colLeft)   (filaTop+1, colLeft+1)
'
'  Limites: filaTop en [1..nR-1],  colLeft en [1..nC-1]
'  Movimientos: arriba / abajo / izquierda / derecha de a 1 celda.
'  No verifica colores intermedios (pinta sobre todo).
'--------------------------------------------------------------------------
Private Function BFSCamino2x2(ByVal sR As Long, ByVal sC As Long, _
                               ByVal tR As Long, ByVal tC As Long, _
                               ByVal nR As Long, ByVal nC As Long) As Variant

    Dim maxR As Long, maxC As Long
    maxR = nR - 1   '  filaTop maxima
    maxC = nC - 1   '  colLeft maxima

    If maxR < 1 Or maxC < 1 Then
        BFSCamino2x2 = Empty
        Exit Function
    End If

    Dim visitado() As Boolean
    Dim padreR()   As Long
    Dim padreC()   As Long
    ReDim visitado(1 To maxR, 1 To maxC)
    ReDim padreR  (1 To maxR, 1 To maxC)
    ReDim padreC  (1 To maxR, 1 To maxC)

    Dim ir As Long, ic As Long
    For ir = 1 To maxR
        For ic = 1 To maxC
            padreR(ir, ic) = 0
            padreC(ir, ic) = 0
        Next ic
    Next ir

    '  Cola BFS
    Dim colaR() As Long, colaC() As Long
    ReDim colaR(1 To maxR * maxC)
    ReDim colaC(1 To maxR * maxC)
    Dim head As Long, tail As Long
    head = 1 : tail = 1

    colaR(tail) = sR
    colaC(tail) = sC
    visitado(sR, sC) = True
    padreR(sR, sC) = -1
    padreC(sR, sC) = -1

    '  4 direcciones
    Dim dr(1 To 4) As Long, dc(1 To 4) As Long
    dr(1) = -1 : dc(1) = 0   ' arriba
    dr(2) =  1 : dc(2) = 0   ' abajo
    dr(3) =  0 : dc(3) = -1  ' izquierda
    dr(4) =  0 : dc(4) =  1  ' derecha

    Dim encontrado As Boolean
    encontrado = (sR = tR And sC = tC)

    Do While head <= tail And Not encontrado
        Dim cr As Long, cc As Long
        cr = colaR(head) : cc = colaC(head)
        head = head + 1

        Dim d As Long
        For d = 1 To 4
            Dim nr2 As Long, nc2 As Long
            nr2 = cr + dr(d)
            nc2 = cc + dc(d)

            If nr2 >= 1 And nr2 <= maxR And nc2 >= 1 And nc2 <= maxC Then
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
        BFSCamino2x2 = Empty
        Exit Function
    End If

    '  Reconstruir camino (destino -> origen) y luego invertir
    Dim tmpR() As Long, tmpC() As Long
    ReDim tmpR(1 To maxR * maxC)
    ReDim tmpC(1 To maxR * maxC)
    Dim n As Long : n = 0
    Dim r As Long, c As Long
    r = tR : c = tC

    Do
        n = n + 1
        tmpR(n) = r
        tmpC(n) = c
        If padreR(r, c) = -1 And padreC(r, c) = -1 Then Exit Do
        Dim auxR As Long, auxC As Long
        auxR = padreR(r, c)
        auxC = padreC(r, c)
        r = auxR
        c = auxC
    Loop

    Dim resultado() As Long
    ReDim resultado(1 To 2, 1 To n)
    Dim j As Long
    For j = 1 To n
        resultado(1, j) = tmpR(n - j + 1)
        resultado(2, j) = tmpC(n - j + 1)
    Next j

    BFSCamino2x2 = resultado
End Function


'--------------------------------------------------------------------------
'  EsBloque2Celdas: TRUE si el rango es exactamente 2 celdas contiguas
'  (1 fila x 2 cols  O  2 filas x 1 col).
'--------------------------------------------------------------------------
Private Function EsBloque2Celdas(rng As Range) As Boolean
    EsBloque2Celdas = (rng.Cells.Count = 2) And _
        ((rng.Rows.Count = 1 And rng.Columns.Count = 2) Or _
         (rng.Rows.Count = 2 And rng.Columns.Count = 1))
End Function


'--------------------------------------------------------------------------
'  EstaEnRango: TRUE si todas las celdas de subRng estan dentro de cont.
'--------------------------------------------------------------------------
Private Function EstaEnRango(ByVal subRng As Range, _
                              ByVal cont As Range) As Boolean
    Dim inter As Range
    On Error Resume Next
    Set inter = Intersect(subRng, cont)
    On Error GoTo 0
    EstaEnRango = (Not inter Is Nothing) And _
                  (inter.Address = subRng.Address)
End Function


'--------------------------------------------------------------------------
'  ObtenerPosBloque: devuelve la esquina superior-izquierda del bloque
'  en coordenadas relativas al area (r1, c1).
'--------------------------------------------------------------------------
Private Sub ObtenerPosBloque(rng As Range, r1 As Long, c1 As Long, _
                              ByRef fila As Long, ByRef col As Long)
    fila = rng.Row    - r1 + 1
    col  = rng.Column - c1 + 1
End Sub


'--------------------------------------------------------------------------
'  HexToColor / ColorToHex  (sin cambios)
'--------------------------------------------------------------------------
Public Function HexToColor(ByVal hex As String) As Long
    Dim s As String
    s = Trim(Replace(hex, "#", ""))
    If Len(s) <> 6 Then
        Err.Raise vbObjectError + 1000, "HexToColor", _
            "Formato invalido: '" & hex & "'. Usa '#RRGGBB'."
    End If
    HexToColor = RGB(CLng("&H" & Mid(s, 1, 2)), _
                     CLng("&H" & Mid(s, 3, 2)), _
                     CLng("&H" & Mid(s, 5, 2)))
End Function

Public Function ColorToHex(ByVal c As Long) As String
    ColorToHex = "#" & Right("00" & Hex(c Mod 256), 2) & _
                       Right("00" & Hex((c \ 256) Mod 256), 2) & _
                       Right("00" & Hex((c \ 65536) Mod 256), 2)
End Function