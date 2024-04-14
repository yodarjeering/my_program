VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Editor 
   Caption         =   "UserForm1"
   ClientHeight    =   4530
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8370.001
   OleObjectBlob   =   "Editor.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub Edditer_Click()
    Call UserForm_Initialize
End Sub

Sub UserForm_Initialize()
    With Me.TextBox1
        .MultiLine = True 'enter　キーで改行を有効にする
        .WordWrap = True
        .EnterKeyBehavior = True
        .Font.Name = "Arial"
        .Font.Size = 12
    End With
    
    'TextBoxに初期値をセット
    Me.TextBox1.Text = "\begin{align}" & vbLf & _
                                   "" & vbLf & _
                                   "\end{align}"
End Sub


Sub RunCommandWithWScriptShell(currentPath As String)
    Dim wsh As Object
    Dim execObject As Object
    Dim strOutput As String
    
    ' WScript.Shellオブジェクトの作成
    Set wsh = CreateObject("WScript.Shell")
    
    ' コマンドプロンプトのコマンドを実行
    Set execObject = wsh.Exec("cmd /c " & currentPath)
    
End Sub

'HTMLに出力する設定を書き出す関数
Private Function BuildTexSettings()
    Dim texScript As String
    
    texScript = "" & _
                    "<head>" & vbLf & _
                    "<meta http-equiv=""X-UA-Compatible"" content=""IE=EmulateIE7"" />" & vbLf & _
                    "<script type=""text/javascript"" src=""https://polyfill.io/v3/polyfill.min.js?features=es6"">" & vbLf & _
                    "  MathJax.Hub.Config({" & vbLf & _
                    "    extensions: [""tex2jax.js""]," & vbLf & _
                    "    jax: [""input/TeX"",""output/HTML-CSS""]," & vbLf & _
                    "    ""HTML-CSS"": {" & vbLf & _
                    "      availableFonts:[]," & vbLf & _
                    "    }" & vbLf & _
                    "  });" & vbLf & _
                    "" & vbLf & _
                    "  window.MathJax = {" & vbLf & _
                    "    tex: {" & vbLf & _
                    "      inlineMath: [['\\(', '\\)'], ['$', '$']],  " & vbLf & _
                    "      displayMath: [['\\[','\\]']]" & vbLf & _
                    "    }," & vbLf & _
                    "    svg: {" & vbLf & _
                    "      fontCache: 'global'" & vbLf & _
                    "    }" & vbLf & _
                    "  };" & vbLf & _
                    "</script>" & vbLf & _
                    "<script id=""MathJax-script"" async src=""https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js""></script>" & vbLf & _
                    "</head>"

    BuildTexSettings = texScript
    
End Function

Sub CommandButton1_Click()
    Dim rowValues() As String
    Dim i As Integer
    
    ' TextBoxのテキストを改行で分割して配列に格納
    rowValues = Split(Me.TextBox1.Text, vbCrLf)
    
    ' 作成するファイルのパスを指定
    Dim currentPath As String
    '作業ディレクトリ直下に描画用ファイルを作成する
    currentPath = ThisWorkbook.Path + "\ShowCalculation.html"
    
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
    
    'コマンドプロンプトを用いて、ShowCalculationを表示する
    Call RunCommandWithWScriptShell(currentPath)
End Sub

