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
    ' �w�肳�ꂽ�͈͂Ƀn�C�p�[�����N��ǉ�
    With rng.Worksheet.Hyperlinks
        .Add Anchor:=rng, Address:=url, TextToDisplay:=IIf(displayText = "", url, displayText)
    End With
End Sub

Function GetCallerCellAddress() As String
    ' �֐����Ăяo�����Z���̃A�h���X��Ԃ�
    Dim caller As Range
    On Error GoTo ErrorHandler
    Set caller = Application.caller
    GetCallerCellAddress = caller.Address
    Exit Function
    
ErrorHandler:
    ' �֐����Z���ȊO����Ăяo���ꂽ�ꍇ�̃G���[�n���h�����O
    GetCallerCellAddress = "Error: Function not called from a cell"
End Function


Function ShowCalculation(inputRange As Range)
    Dim cell As Range
    Dim rowValues() As String
    Dim i As Integer
    
    
    ' �z��̃T�C�Y��������
    ReDim rowValues(inputRange.Rows.Count - 1)
    
    i = 0
    ' �͈͓��̊e�s�ɂ��ă��[�v
    For Each cell In inputRange.Rows
        ' �e�s�̍ŏ��̃Z���̒l�𕶎���Ƃ��Ĕz��Ɋi�[
        rowValues(i) = CStr(cell.Cells(1, 1).Value)
        i = i + 1
    Next cell

     ' �쐬����t�@�C���̃p�X���w��
    Dim currentPath As String
    currentPath = ThisWorkbook.Path + "\test.html"
    Dim fileNumber As Integer
    fileNumber = FreeFile
    'HTML�ɏo�͂���l
    Dim output As String
    output = BuildTexSettings & vbCrLf

    ' �t�@�C�����J���ď������݃��[�h�ŊJ��
    Open currentPath For Output As #fileNumber
    ' �t�@�C���Ƀe�L�X�g����������
    Print #fileNumber, output

    
    For i = 0 To UBound(rowValues)
        Print #fileNumber, rowValues(i)
   Next i

    ' �t�@�C�������
    Close #fileNumber
    
    ShowCalculation = currentPath
    

End Function
