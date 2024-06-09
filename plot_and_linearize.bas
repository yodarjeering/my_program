
Sub PlotDataWithTrendline(sheetName As String, dataRange As String)
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As Chart
    Dim rng As Range

    ' シートを設定
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' データ範囲を設定
    Set rng = ws.Range(dataRange)

    ' グラフオブジェクトを追加
    Set chartObj = ws.ChartObjects.Add(Left:=100, Width:=375, Top:=50, Height:=225)
    Set chart = chartObj.Chart

    ' グラフのデータソースを設定
    chart.SetSourceData Source:=rng

    ' グラフの種類を散布図に設定
    chart.ChartType = xlXYScatterLines

    ' 近似曲線を追加
    With chart.SeriesCollection(1)
        .Trendlines.Add Type:=xlLinear, Forward:=0, Backward:=0, DisplayEquation:=True, DisplayRSquared:=True
    End With

    ' グラフのタイトルと軸ラベルを設定
    chart.HasTitle = True
    chart.ChartTitle.Text = "データと近似曲線"
    chart.Axes(xlCategory, xlPrimary).HasTitle = True
    chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "X軸"
    chart.Axes(xlValue, xlPrimary).HasTitle = True
    chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Y軸"
End Sub

Function RangeToArray(range As Range) As Variant
    Dim arr() As Variant
    arr = range.Value

    

    RangeToArray = arr
End Function


