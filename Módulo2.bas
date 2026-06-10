Attribute VB_Name = "M�dulo2"
Sub OcultarBARRAS()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", false)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False
    ActiveSheet.Shapes("Mostrar Barras").Visible = True
    ActiveSheet.Shapes("ocultar Barras").Visible = False
End Sub
' Mostrar las barras y elementos de la interfaz
Sub MostarBARRAS()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", true)"
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayHeadings = true
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveSheet.Shapes("ocultar Barras").Visible = True
    ActiveSheet.Shapes("Mostrar Barras").Visible = False
End Sub

Public Sub Envolventes_GFLX()
    Dim shp As Shape
    Dim groupname As String
    groupname = "ENVOLVENTESGFLX"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
         ' Oculta los dem�s grupos
        ActiveSheet.Shapes("BARRAS_1").Visible = False
        activeSheet.Shapes("ALMACEN_GFLX").Visible = False
    End If
End Sub
Public Sub BARRAS_GFLX()
    Dim shp As Shape
    Dim groupname As String
    groupname = "BARRAS_1"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
         ' Oculta los dem�s grupos
        ActiveSheet.Shapes("ENVOLVENTESGFLX").Visible = False
        activeSheet.Shapes("ALMACEN_GFLX").Visible = False
    End If
End Sub
Public sub Almacen_GFLX()
    Dim shp As Shape
    Dim groupname As String
    groupname = "ALMACEN_GFLX"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
         ' Oculta los dem�s grupos
        ActiveSheet.Shapes("BARRAS_1").Visible = False
        ActiveSheet.Shapes("ENVOLVENTESGFLX").Visible = False
    End If
End Sub

'=========================
'RESTABLECER CAJETINES
'=========================
Public sub cajetin_GFLX()
    thisworkbook.sheets("frontal GFX").range("BK43:IM145").clear
end sub
