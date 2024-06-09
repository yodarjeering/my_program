Sub CreateChartFromSelection()
    Dim selectedRange As Range
    Dim chart As Chart
    
    ' 選択された範囲を取得
    Set selectedRange = Selection
    
    ' 新しいチャートを作成
    Set chart = Charts.Add
    With chart
        .SetSourceData Source:=selectedRange
        .ChartType = xlLine ' グラフの種類を設定（例：線グラフ）
        .Location Where:=xlLocationAsObject, Name:="Sheet1"
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