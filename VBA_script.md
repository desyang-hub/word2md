## 使用VBA宏，将所有MathType公式转换为TeX格式

### 进入word, 使用Alt+F11 打开VBA编辑器，插入模块，输入以下代码，保存为宏，运行宏即可
```
Sub ConvertAllMathTypeToTeX()
    Dim i As Integer
    Dim sh As InlineShape
    Dim count As Integer
    
    count = 0
    For i = ActiveDocument.InlineShapes.count To 1 Step -1
        Set sh = ActiveDocument.InlineShapes(i)
        If sh.Type = wdInlineShapeEmbeddedOLEObject Then
            sh.Select
            ' 调用 MathType 的 TeXToggle（就是 Alt+\ 绑定的命令）
            On Error Resume Next
            Application.Run "MTCommand_TeXToggle"
            On Error GoTo 0
            count = count + 1
        End If
    Next i
    
    MsgBox "转换完成，共 " & count & " 个公式"
End Sub
```