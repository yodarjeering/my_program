' 　行いたい動作の流れ

' For fileName in FileNameCnt
'     For Csvファイルの列

'         選択された列のプロット
'         Plot()
'         値の判定
'         Judge()



Sub SelectCsv()
    Dim fd As FileDialog
    Dim fileName As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Input")

    ' ファイル選択ダイアログの設定
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Filters.Clear
        .Filters.Add "CSV Files",".csv"
        If .Show = True Then
            fileName = .SelectedItems(1)
        Else
            MsgBox "ファイルは選択されませんでした"
            Exit Sub
        End If
    End With

End Sub

Sub ImportCSVtoWorksheet(filePath As String, sheetName As String)
    Dim ws As Worksheet
    Dim qt As QueryTable
    
    ' 新しいワークシートを追加
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName
    
    ' CSVファイルをQueryTableを使ってインポート
    Set qt = ws.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=ws.Range("A1"))
    With qt
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .Refresh
    End With
    qt.Delete  ' インポート後はQueryTableオブジェクトを削除
End Sub


Function SheetDataToArray(ws As Worksheet) As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range
    
    ' 使用されている範囲の最終行と最終列を取得
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' データ範囲を設定
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' データ範囲を配列に読み込む
    SheetDataToArray = dataRange.Value
End Function

'実行例
Sub TestImportAndConvert()
    Dim ws As Worksheet
    Dim data As Variant
    
    ' CSVをインポート
    Call ImportCSVtoWorksheet("C:\path\to\yourfile.csv", "ImportedCSV")
    
    ' インポートされたワークシートを取得
    Set ws = ThisWorkbook.Sheets("ImportedCSV")
    
    ' ワークシートのデータを配列に変換
    data = SheetDataToArray(ws)
    
    ' 配列の内容をデバッグ出力（例）
    Debug.Print data(1, 1)  ' 配列の最初の要素を出力
End Sub




Sub button1_click()
        Call SelectCsv
        

End Sub

Sub SelectCsv()
    Dim fd As FileDialog
    Dim fileName As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Input")

    ' ファイル選択ダイアログの設定
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        If .Show = True Then
            fileName = .SelectedItems(1)
        Else
            MsgBox "ファイルは選択されませんでした"
            Exit Sub
        End If
    End With
    
    Call SelectColumnAndPlot(fileName)
    

End Sub

' Debug 用のcsvを読み取る関数
Function LoadCSV() As Variant
    Dim fileNum As Integer
    Dim rowData As String
    Dim data As Variant
    Dim dataList As Collection
    Dim i As Long
    Dim filePath As String  ' ファイルパス保存用
    filePath = "C:\Users\Public\ForCalc\PlotChart\COVID-19_Outcomes_by_Vaccination_Status.csv"
    

    fileNum = FreeFile
    Open filePath For Input As #fileNum

    Set dataList = New Collection
    
    Do Until EOF(fileNum)
        Line Input #fileNum, rowData

        dataList.Add Split(rowData, ",")  ' カンマで区切って配列に変換し、コレクションに追加

    Loop

    Close #fileNum

    ' コレクションを配列に変換
    ReDim data(1 To dataList.Count)
        
    For i = 1 To dataList.Count
        data(i) = dataList(i)
    Next i
    LoadCSV = data
    
End Function


Sub PlotData(data As Variant)
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim i As Long
    Dim Left, Width, Top, Height As Integer
    Set ws = ThisWorkbook.Sheets("Plot")
    Dim selectedColumn As Integer
    selectedColumn = 16
    Left = 100
    Width = 375
    Top = 50
    Height = 225


    ' 散布図を追加
    Set chartObj = ws.ChartObjects.Add(Left:=Left, Width:=Width, Top:=Top + Height, Height:=Height)
    With chartObj.Chart
        .ChartType = xlXYScatter
        
        ' 仮にX値の配列を用意
        ReDim XValues(1 To UBound(data) - LBound(data) + 1)
        ReDim yValues(1 To UBound(data) - LBound(data) + 1)
        
        ' データを配列に設定
        For i = LBound(data) + 1 To UBound(data)
            XValues(i - LBound(data) + 1) = i
            yValues(i - LBound(data) + 1) = CInt(data(i)(selectedColumn))
            
        Next i
        
        ' 新しい系列を追加し、配列全体を設定
        .SeriesCollection.NewSeries
        .SeriesCollection(1).XValues = XValues
        .SeriesCollection(1).Values = yValues
        '適宜ラベルなど追加する
    End With
    
    
    
End Sub

Sub DebugFunc()
    Dim data As Variant
    data = LoadCSV()
    Call PlotData(data)
End Sub
