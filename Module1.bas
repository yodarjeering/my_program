Attribute VB_Name = "Module1"


Function BuildTexSettings()
    Dim texScript As String
    
    texScript = "" & _
    "<head>" & _
    "<meta http-equiv=""X-UA-Compatible"" content=""IE=EmulateIE7"" />" & _
    "<script type=""text/javascript"" src=""https://polyfill.io/v3/polyfill.min.js?features=es6"">" & _
    "  MathJax.Hub.Config({" & _
    "    extensions: [""tex2jax.js""]," & _
    "    jax: [""input/TeX"",""output/HTML-CSS""]," & _
    "    ""HTML-CSS"": {" & _
    "      availableFonts:[]," & _
    "    }" & _
    "  });" & _
    "" & _
    "  window.MathJax = {" & _
    "    tex: {" & _
    "      inlineMath: [['\\(','\\)']]," & _
    "      displayMath: [['\\[','\\]']]" & _
    "    }," & _
    "    svg: {" & _
    "      fontCache: 'global'" & _
    "    }" & _
    "  };" & _
    "</script>" & _
    "<script id=""MathJax-script"" async src=""https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js""></script>" & _
    "</head>"

    BuildTexSettings = texScript
    
End Function

Sub AddHyperlink(rng As Range, url As String, Optional displayText As String = "")
    ' 指定された範囲にハイパーリンクを追加
    With rng.Worksheet.Hyperlinks
        .Add Anchor:=rng, Address:=url, TextToDisplay:=IIf(displayText = "", url, displayText)
    End With
End Sub

Function GetCallerCellAddress() As String
    ' 関数を呼び出したセルのアドレスを返す
    Dim caller As Range
    On Error GoTo ErrorHandler
    Set caller = Application.caller
    GetCallerCellAddress = caller.Address
    Exit Function
    
ErrorHandler:
    ' 関数がセル以外から呼び出された場合のエラーハンドリング
    GetCallerCellAddress = "Error: Function not called from a cell"
End Function


Function ShowCalculation(inputRange As Range)
    Dim cell As Range
    Dim rowValues() As String
    Dim i As Integer
    
    
    ' 配列のサイズを初期化
    ReDim rowValues(inputRange.Rows.Count - 1)
    
    i = 0
    ' 範囲内の各行についてループ
    For Each cell In inputRange.Rows
        ' 各行の最初のセルの値を文字列として配列に格納
        rowValues(i) = CStr(cell.Cells(1, 1).Value)
        i = i + 1
    Next cell

     ' 作成するファイルのパスを指定
    Dim currentPath As String
    currentPath = ThisWorkbook.Path + "\test.html"
    Dim fileNumber As Integer
    fileNumber = FreeFile
    'HTMLに出力する値
    Dim output As String
    output = BuildTexSettings & vbCrLf

    ' ファイルを開いて書き込みモードで開く
    Open currentPath For Output As #fileNumber
    ' ファイルにテキストを書き込む
    Print #fileNumber, output

    
    For i = 0 To UBound(rowValues)
        Print #fileNumber, rowValues(i)
   Next i

    ' ファイルを閉じる
    Close #fileNumber
    
    ShowCalculation = currentPath
    

End Function
