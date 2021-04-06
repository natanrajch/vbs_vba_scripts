
Sub PosicionLineaPromedio()
' It modifies a custom line inside the charts of an excel sheet that adds mean-so-far info to these charts
'Its hard coded as hell, but it doesnt need that much flexibility either so be it
' Modifica gráfico a gráfico la posición de la línea del promedio acumulado del mes de FM

Dim MediaEscalerasFM As Double
Dim MediaEscalerasNoFM As Double
Dim MediaAscensoresFM As Double
Dim Subir As Double

Application.ScreenUpdating = False

'Modifica altura de línea FM TOTALES ESCALERAS
Worksheets("Acumulado a hoy").Activate
MediaEscalerasFM = Cells(12, "G").Value
Worksheets("Total").Activate
    ActiveSheet.ChartObjects("1 Gráfico").Activate
    ActiveChart.Shapes.Range(Array("1 Conector recto")).Select
    Selection.ShapeRange.IncrementTop 1208.5714173228
    Selection.ShapeRange.IncrementTop -45
    'Mueve la línea hasta abajo de todo y luego la sube hasta el nivel de 90%
    Subir = ((MediaEscalerasFM - 0.9) / 0.1) * (-195)
'Hace un proporcional entre la altura total (-195) y el % acumulado, siendo 95% = 50% de la altura que debe subir
    Selection.ShapeRange.IncrementTop (Subir)

'Modifica el texto del cuadro
    ActiveChart.Shapes.Range(Array("3 CuadroTexto")).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = _
        "Promedio acumulado del mes C/FM: %" & Round(MediaEscalerasFM * 100, 2)

'Modifica altura de línea Sin FM TOTALES ESCALERAS
Worksheets("Acumulado a hoy").Activate
MediaEscalerasNoFM = Cells(12, "H").Value
Worksheets("Total").Activate
    ActiveSheet.ChartObjects("1 Gráfico").Activate
    ActiveChart.Shapes.Range(Array("5 Conector recto")).Select
    Selection.ShapeRange.IncrementTop 1208.5714173228
    Selection.ShapeRange.IncrementTop -45
    'Mueve la línea hasta abajo de todo y luego la sube hasta el nivel de 90%
    
Subir = ((MediaEscalerasNoFM - 0.9) / 0.1) * (-195)
'Hace un proporcional entre la altura total (-187) y el % acumulado, siendo 95% = 50% de la altura que debe subir
    Selection.ShapeRange.IncrementTop (Subir)

'Modifica altura de línea FM TOTALES ASCENSORES
Worksheets("Acumulado a hoy").Activate
MediaAscensoresFM = Cells(24, "G").Value
Worksheets("Total").Activate
    ActiveSheet.ChartObjects("2 Gráfico").Activate
    ActiveChart.Shapes.Range(Array("1 Conector recto")).Select
    Selection.ShapeRange.IncrementTop 1208.5714173228
    Selection.ShapeRange.IncrementTop -43.9285826772
    'Mueve la línea hasta abajo de todo y luego la sube hasta el nivel de 90%
    Subir = ((MediaAscensoresFM - 0.7) / 0.3) * (-190)
'Hace un proporcional entre la altura total (-187) y el % acumulado, siendo 85% = 50% de la altura que debe subir
    Selection.ShapeRange.IncrementTop (Subir)

'Modifica el texto del cuadro
    ActiveChart.Shapes.Range(Array("3 CuadroTexto")).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = _
        "Promedio acumulado del mes C/FM: %" & Round(MediaAscensoresFM * 100, 2)

'Modifica altura de línea FM ESCALERAS - 1
Worksheets("Acumulado a hoy").Activate
MediaEscalerasFM = Cells(6, "G").Value
Worksheets("Línea a línea").Activate
    ActiveSheet.ChartObjects("1 Gráfico").Activate
    ActiveChart.Shapes.Range(Array("5 Conector recto")).Select
    Selection.ShapeRange.IncrementTop 1208.5714173228
    Selection.ShapeRange.IncrementTop -44.25
    'Mueve la línea hasta abajo de todo y luego la sube hasta el nivel de 90%
    Subir = ((MediaEscalerasFM - 0.9) / 0.1) * (-188.25)
'Hace un proporcional entre la altura total (-187) y el % acumulado, siendo 95% = 50% de la altura que debe subir
    Selection.ShapeRange.IncrementTop (Subir)

'Modifica altura de línea FM ESCALERAS - 2
Worksheets("Acumulado a hoy").Activate
MediaEscalerasFM = Cells(7, "G").Value
Worksheets("Línea a línea").Activate
    ActiveSheet.ChartObjects("2 Gráfico").Activate
    ActiveChart.Shapes.Range(Array("6 Conector recto")).Select
    Selection.ShapeRange.IncrementTop 1208.5714173228
    Selection.ShapeRange.IncrementTop -44.25
    'Mueve la línea hasta abajo de todo y luego la sube hasta el nivel de 90%
    Subir = ((MediaEscalerasFM - 0.9) / 0.1) * (-188.25)
'Hace un proporcional entre la altura total (-187) y el % acumulado, siendo 95% = 50% de la altura que debe subir
    Selection.ShapeRange.IncrementTop (Subir)

'Modifica altura de línea FM ESCALERAS - 3
Worksheets("Acumulado a hoy").Activate
MediaEscalerasFM = Cells(8, "G").Value
Worksheets("Línea a línea").Activate
    ActiveSheet.ChartObjects("3 Gráfico").Activate
    ActiveChart.Shapes.Range(Array("4 Conector recto")).Select
    Selection.ShapeRange.IncrementTop 1208.5714173228
    Selection.ShapeRange.IncrementTop -44.25
    'Mueve la línea hasta abajo de todo y luego la sube hasta el nivel de 90%
    Subir = ((MediaEscalerasFM - 0.85) / 0.15) * (-188.25)
'Hace un proporcional entre la altura total (-187) y el % acumulado, siendo 92.5% = 50% de la altura que debe subir
    Selection.ShapeRange.IncrementTop (Subir)

'Modifica altura de línea FM ESCALERAS - 4
Worksheets("Acumulado a hoy").Activate
MediaEscalerasFM = Cells(9, "G").Value
Worksheets("Línea a línea").Activate
    ActiveSheet.ChartObjects("4 Gráfico").Activate
    ActiveChart.Shapes.Range(Array("4 Conector recto")).Select
    Selection.ShapeRange.IncrementTop 1208.5714173228
    Selection.ShapeRange.IncrementTop -44.25
    'Mueve la línea hasta abajo de todo y luego la sube hasta el nivel de 90%
    Subir = ((MediaEscalerasFM - 0.9) / 0.1) * (-188.25)
'Hace un proporcional entre la altura total (-187) y el % acumulado, siendo 95% = 50% de la altura que debe subir
    Selection.ShapeRange.IncrementTop (Subir)

'Modifica altura de línea FM ESCALERAS - 5
Worksheets("Acumulado a hoy").Activate
MediaEscalerasFM = Cells(10, "G").Value
Worksheets("Línea a línea").Activate
    ActiveSheet.ChartObjects("5 Gráfico").Activate
    ActiveChart.Shapes.Range(Array("5 Conector recto")).Select
    Selection.ShapeRange.IncrementTop 1208.5714173228
    Selection.ShapeRange.IncrementTop -44.25
    'Mueve la línea hasta abajo de todo y luego la sube hasta el nivel de 90%
    Subir = ((MediaEscalerasFM - 0.8) / 0.2) * (-188.25)
'Hace un proporcional entre la altura total (-187) y el % acumulado, siendo 90% = 50% de la altura que debe subir
    Selection.ShapeRange.IncrementTop (Subir)


'Modifica altura de línea FM ESCALERAS - 6
Worksheets("Acumulado a hoy").Activate
MediaEscalerasFM = Cells(11, "G").Value
Worksheets("Línea a línea").Activate
    ActiveSheet.ChartObjects("6 Gráfico").Activate
    ActiveChart.Shapes.Range(Array("4 Conector recto")).Select
    Selection.ShapeRange.IncrementTop 1208.5714173228
    Selection.ShapeRange.IncrementTop -44.25
    'Mueve la línea hasta abajo de todo y luego la sube hasta el nivel de 90%
    Subir = ((MediaEscalerasFM - 0.9) / 0.1) * (-188.25)
'Hace un proporcional entre la altura total (-187) y el % acumulado, siendo 95% = 50% de la altura que debe subir
    Selection.ShapeRange.IncrementTop (Subir)



'Modifica altura de línea FM ASCENSORES - 1
Worksheets("Acumulado a hoy").Activate
MediaAscensoresFM = Cells(18, "G").Value
Worksheets("Línea a línea").Activate
    ActiveSheet.ChartObjects("7 Gráfico").Activate
    ActiveChart.Shapes.Range(Array("4 Conector recto")).Select
    Selection.ShapeRange.IncrementTop 1208.5714173228
    Selection.ShapeRange.IncrementTop -44.25
    'Mueve la línea hasta abajo de todo y luego la sube hasta el nivel de 90%
    Subir = ((MediaAscensoresFM - 0.4) / 0.6) * (-188.25)
'Hace un proporcional entre la altura total (-187) y el % acumulado, siendo 70% = 50% de la altura que debe subir
    Selection.ShapeRange.IncrementTop (Subir)

'Modifica altura de línea FM ASCENSORES - 2
Worksheets("Acumulado a hoy").Activate
MediaAscensoresFM = Cells(19, "G").Value
Worksheets("Línea a línea").Activate
    ActiveSheet.ChartObjects("8 Gráfico").Activate
    ActiveChart.Shapes.Range(Array("4 Conector recto")).Select
    Selection.ShapeRange.IncrementTop 1208.5714173228
    Selection.ShapeRange.IncrementTop -44.25
    'Mueve la línea hasta abajo de todo y luego la sube hasta el nivel de 90%
Subir = ((MediaAscensoresFM - 0.4) / 0.6) * (-188.25)
'Hace un proporcional entre la altura total (-187) y el % acumulado, siendo 70% = 50% de la altura que debe subir
    Selection.ShapeRange.IncrementTop (Subir)

'Modifica altura de línea FM ASCENSORES - 3
Worksheets("Acumulado a hoy").Activate
MediaAscensoresFM = Cells(20, "G").Value
Worksheets("Línea a línea").Activate
    ActiveSheet.ChartObjects("9 Gráfico").Activate
    ActiveChart.Shapes.Range(Array("4 Conector recto")).Select
    Selection.ShapeRange.IncrementTop 1208.5714173228
    Selection.ShapeRange.IncrementTop -44.25
    'Mueve la línea hasta abajo de todo y luego la sube hasta el nivel de 90%
Subir = ((MediaAscensoresFM - 0.4) / 0.6) * (-188.25)
'Hace un proporcional entre la altura total (-187) y el % acumulado, siendo 70% = 50% de la altura que debe subir
    Selection.ShapeRange.IncrementTop (Subir)


'Modifica altura de línea FM ASCENSORES - 4
Worksheets("Acumulado a hoy").Activate
MediaAscensoresFM = Cells(21, "G").Value
Worksheets("Línea a línea").Activate
    ActiveSheet.ChartObjects("10 Gráfico").Activate
    ActiveChart.Shapes.Range(Array("4 Conector recto")).Select
    Selection.ShapeRange.IncrementTop 1208.5714173228
    Selection.ShapeRange.IncrementTop -44.25
    'Mueve la línea hasta abajo de todo y luego la sube hasta el nivel de 90%
    Subir = ((MediaAscensoresFM - 0.4) / 0.6) * (-188.25)
'Hace un proporcional entre la altura total (-187) y el % acumulado, siendo 70% = 50% de la altura que debe subir
    Selection.ShapeRange.IncrementTop (Subir)

'Modifica altura de línea FM ASCENSORES - 5
Worksheets("Acumulado a hoy").Activate
MediaAscensoresFM = Cells(22, "G").Value
Worksheets("Línea a línea").Activate
    ActiveSheet.ChartObjects("11 Gráfico").Activate
    ActiveChart.Shapes.Range(Array("1 Conector recto")).Select
    Selection.ShapeRange.IncrementTop 1208.5714173228
    Selection.ShapeRange.IncrementTop -44.25
    'Mueve la línea hasta abajo de todo y luego la sube hasta el nivel de 90%
    Subir = ((MediaAscensoresFM - 0.4) / 0.6) * (-188.25)
'Hace un proporcional entre la altura total (-187) y el % acumulado, siendo 70% = 50% de la altura que debe subir
    Selection.ShapeRange.IncrementTop (Subir)



'Modifica altura de línea FM ASCENSORES - 6
Worksheets("Acumulado a hoy").Activate
MediaAscensoresFM = Cells(23, "G").Value
Worksheets("Línea a línea").Activate
    ActiveSheet.ChartObjects("12 Gráfico").Activate
    ActiveChart.Shapes.Range(Array("4 Conector recto")).Select
    Selection.ShapeRange.IncrementTop 1208.5714173228
    Selection.ShapeRange.IncrementTop -44.25
    'Mueve la línea hasta abajo de todo y luego la sube hasta el nivel de 90%
    Subir = ((MediaAscensoresFM - 0.4) / 0.6) * (-188.25)
'Hace un proporcional entre la altura total (-187) y el % acumulado, siendo 70% = 50% de la altura que debe subir
    Selection.ShapeRange.IncrementTop (Subir)

Worksheets(1).Activate
Application.ScreenUpdating = True



End Sub
Sub CambioSemana2()
'This changes a field inside pivot tables

Dim pt As PivotTable
Dim nuevaSemana As String

Application.ScreenUpdating = False


Worksheets(1).Activate
nuevaSemana = Cells(1, "N").Value

For i = 1 To 2

For Each pt In Worksheets(i).PivotTables
pt.PivotFields("Semana").CurrentPage = nuevaSemana
Next pt
Next

Application.ScreenUpdating = True

End Sub


Sub ExportChart()
'Exports dynamic-chart to png file
 

    Dim objChrt As ChartObject
    Dim myChart As Chart
    Dim i As Integer
    Dim x As Integer
    Dim dViernesInforme As Double
    
    dViernesInforme = Date - Weekday(Date, vbFriday) + 1
    
    'The export quality depends on the zoom level. Whatyougonnado.
    Application.ScreenUpdating = False
    ThisWorkbook.Worksheets(1).Activate
    ActiveWindow.Zoom = 190
    ThisWorkbook.Worksheets(2).Activate
    ActiveWindow.Zoom = 190
    Application.ScreenUpdating = True

    For i = 1 To 2
    x = 1
    
    For Each objChrt In Sheets(i).ChartObjects
'    Set objChrt = Sheets(i).ChartObjects(3)
    
    objChrt.Activate 'Línea agregada por el bug donde a veces y random exporta archivos vacíos. GRacia stakoverflou
    
    Set myChart = objChrt.Chart

    myFileName = "Graf-" & i & "-" & x & "-" & Day(dViernesInforme) & "-" & Month(dViernesInforme) & "-" & Year(dViernesInforme) & ".png"

    On Error Resume Next
    Kill ThisWorkbook.Path & "\Gráficos\" & myFileName
    On Error GoTo 0
    
    
    
    myChart.Export Filename:=ThisWorkbook.Path & "\Gráficos\" & myFileName, Filtername:="PNG"
    
    x = x + 1
    Next objChrt
    Next i

    Application.ScreenUpdating = False
    ThisWorkbook.Worksheets(1).Activate
    ActiveWindow.Zoom = 80
    ThisWorkbook.Worksheets(2).Activate
    ActiveWindow.Zoom = 80
    Application.ScreenUpdating = True

ThisWorkbook.Save
ThisWorkbook.Close (False)
'    MsgBox "OK"
End Sub

'EVENT LISTENER - Poner dentro de worksheet, no de un modulo. 
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCells As Range

    ' The variable KeyCells contains the cells that will
    ' cause an alert when they are changed.
    Set KeyCells = Range("N1")
    
    If Not Application.Intersect(KeyCells, Range(Target.Address)) _
           Is Nothing Then

        ' Display a message when one of the designated cells has been
        ' changed.
        ' Place your code here.
        Call CambioSemana2
       
    End If
End Sub