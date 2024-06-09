Sub CreateChartFromSelection()
    Dim selectedRange As Range
    Dim chart As chart
    
    ' 選択された範囲を取得
    Set selectedRange = Selection
    
    ' 新しいチャートを作成
    Set chart = Charts.Add
    With chart
        .SetSourceData Source:=selectedRange
        .chartType = xlLine ' グラフの種類を設定（例：線グラフ）
        .Location Where:=xlLocationAsObject, Name:=ActiveSheet.Name
    End With
    
End Sub


Sub CreateTableFromSelection()
Dim selectedRange As Range
Dim listObj As ListObject

' 選択された範囲を取得
Set selectedRange = Selection

' 選択範囲にテーブルを作成
Set listObj = ActiveSheet.ListObjects.Add(xlSrcRange, selectedRange, , xlYes)
With listObj
    .Name = "MyTable"
    .TableStyle = "TableStyleLight9" ' テーブルのスタイルを設定
End With

End Sub
    
Sub CreateApproximationCurve()
    Dim selectedRange As Range
    Dim chart As chart
    Dim series As series
    Dim trendline As trendline
    Dim chartType As Integer
    Dim coefficients As String
    Dim outputCell As Range
    
    ' 選択された範囲を取得
    Set selectedRange = Selection
    
    ' 新しいチャートを作成
    Set chart = ActiveSheet.Shapes.AddChart2(251, xlXYScatter).chart
    
    With chart
        .SetSourceData Source:=selectedRange
        .chartType = xlXYScatter ' 散布図を基本とする
        
        ' チャートをアクティブシートに埋め込む
        .Location Where:=xlLocationAsObject, Name:=ActiveSheet.Name
    End With
    
    ' 近似曲線の種類を選択するためのダイアログボックスを表示
    chartType = Application.InputBox("近似曲線の種類を入力してください (1: 線形, 2: 指数, 3: 対数, 4: 多項式)", Type:=1)
    
    ' シリーズを取得
    Set series = chart.SeriesCollection(1)
    
    ' 近似曲線の種類に応じて設定
    Select Case chartType
        Case 1
            Set trendline = series.Trendlines.Add(Type:=xlLinear)
        Case 2
            Set trendline = series.Trendlines.Add(Type:=xlExponential)
        Case 3
            Set trendline = series.Trendlines.Add(Type:=xlLogarithmic)
        Case 4
            Set trendline = series.Trendlines.Add(Type:=xlPolynomial, Order:=2) ' 2次多項式
        Case Else
            MsgBox "無効な入力です。線形近似を使用します。"
            Set trendline = series.Trendlines.Add(Type:=xlLinear)
    End Select
    
    ' 近似曲線の数式とR^2値の表示状態を確認
    Debug.Print "Equation displayed: " & trendline.DisplayEquation
    Debug.Print "R-squared displayed: " & trendline.DisplayRSquared

    ' 係数のテキストを出力
    'Debug.Print "Coefficients text: " & trendline.DataLabel.Text

    ' グラフを更新
    chart.Refresh

    ' 更新後の係数のテキストを再度出力
    'Debug.Print "Updated coefficients text: " & trendline.DataLabel.Text
    
    ' 近似曲線の数式をグラフ中に表示
    trendline.DisplayEquation = True
    ' R^2値
    trendline.DisplayRSquared = True
    
    ' グラフの更新を強制
    chart.Refresh
    
    ' 係数をシートの任意の場所に表示
    coefficients = trendline.DataLabel.Text
    Set outputCell = Application.InputBox("係数を表示するセルを選択してください", Type:=8)
    outputCell.Value = coefficients
End Sub

Sub CalculateAndDisplayCoefficients()
    Dim selectedRange As Range
    Dim chart As chart
    Dim series As series
    Dim trendline As trendline
    Dim outputCell As Range
    Dim coefficients As String
    
    ' 選択された範囲を取得
    Set selectedRange = Selection
    
    ' 選択範囲が正しい形式か確認
    If selectedRange.Columns.Count <> 2 Then
        MsgBox "選択範囲は2列である必要があります。"
        Exit Sub
    End If
    
    ' 新しいチャートを作成
    Set chart = ActiveSheet.Shapes.AddChart2(251, xlXYScatter).chart
    With chart
        .SetSourceData Source:=selectedRange
        .chartType = xlXYScatter
        .Location Where:=xlLocationAsObject, Name:=ActiveSheet.Name
    End With
    
    ' シリーズを取得
    Set series = chart.SeriesCollection(1)
    
    ' 線形近似曲線を追加
    Set trendline = series.Trendlines.Add(Type:=xlLinear)
    
    ' 近似曲線の数式??? R^2値を表示
    trendline.DisplayEquation = True
    trendline.DisplayRSquared = True
    
    ' グラフの更新を強制
    chart.Refresh
    
    ' 係数を取得
    coefficients = trendline.DataLabel.Text
    
    ' 係数を表示するセルをユーザーに選択させる
    Set outputCell = Application.InputBox("係数を表示するセルを選択してください", Type:=8)
    
    ' 係数をセルに設定
    outputCell.Value = coefficients
    
    ' 作成したチャートを削除
    chart.Parent.Delete
End Sub
