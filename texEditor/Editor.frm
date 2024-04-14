VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Editor 
   Caption         =   "UserForm1"
   ClientHeight    =   4530
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8370.001
   OleObjectBlob   =   "Editor.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
        .MultiLine = True 'enter�@�L�[�ŉ��s��L���ɂ���
        .WordWrap = True
        .EnterKeyBehavior = True
        .Font.Name = "Arial"
        .Font.Size = 12
    End With
    
    'TextBox�ɏ����l���Z�b�g
    Me.TextBox1.Text = "\begin{align}" & vbLf & _
                                   "" & vbLf & _
                                   "\end{align}"
End Sub


Sub RunCommandWithWScriptShell(currentPath As String)
    Dim wsh As Object
    Dim execObject As Object
    Dim strOutput As String
    
    ' WScript.Shell�I�u�W�F�N�g�̍쐬
    Set wsh = CreateObject("WScript.Shell")
    
    ' �R�}���h�v�����v�g�̃R�}���h�����s
    Set execObject = wsh.Exec("cmd /c " & currentPath)
    
End Sub

'HTML�ɏo�͂���ݒ�������o���֐�
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
    
    ' TextBox�̃e�L�X�g�����s�ŕ������Ĕz��Ɋi�[
    rowValues = Split(Me.TextBox1.Text, vbCrLf)
    
    ' �쐬����t�@�C���̃p�X���w��
    Dim currentPath As String
    '��ƃf�B���N�g�������ɕ`��p�t�@�C�����쐬����
    currentPath = ThisWorkbook.Path + "\ShowCalculation.html"
    
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
    
    '�R�}���h�v�����v�g��p���āAShowCalculation��\������
    Call RunCommandWithWScriptShell(currentPath)
End Sub

