

Sub 调整指定ppt()
    ' 定义变量
    Dim pptFilePath As String
    Dim presentation As presentation
      
    ' 设置 pptx 文件路径
    pptFilePath = "/Users/xxx.pptx"
      
    ' 打开 pptx 文件
    Set presentation = Presentations.Open(pptFilePath)
    
        复制每页幻灯片
        批量移动shapes
        '删除多余形状
        
        MsgBox "处理完毕！"
        
End Sub



Sub 形状位置()

t = ActiveWindow.Selection.ShapeRange.Top
l = ActiveWindow.Selection.ShapeRange.Left
Debug.Print "top：" & t, "left：" & l

w = ActiveWindow.Selection.ShapeRange.Width
h = ActiveWindow.Selection.ShapeRange.Height
Debug.Print "width：" & w, "height：" & h

a = ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Paragraphs(1).Text
Debug.Print "文字：" & a

c = Replace(ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Paragraphs(1).Text, vbCr, "")
Debug.Print "去除段落标记后的文字：" & c

b = ActiveWindow.Selection.ShapeRange.Type
Debug.Print "类型：" & b

End Sub


Sub 复制每页幻灯片()

sn = Application.ActivePresentation.Slides.Count
Debug.Print sn

For i = 1 To sn * 2 Step 2

    Application.ActivePresentation.Slides(i).Duplicate

Next

'MsgBox "复制处理完毕！"

End Sub



Sub 批量移动shapes()

sn = Application.ActivePresentation.Slides.Count
Debug.Print sn

For i = 1 To sn

    Set myDocument = Application.ActivePresentation.Slides(i)
    rn = myDocument.Shapes.Count
    Debug.Print rn
    
    ActivePresentation.Slides(i).Select
    ActivePresentation.Slides(i).Shapes.SelectAll
    
    With ActiveWindow.Selection.ShapeRange

        ' 奇数页
        If i Mod 2 = 1 Then
            .IncrementLeft -240
            .IncrementTop -135
        Else
        '偶数页
            .IncrementLeft -240
            .IncrementTop -135
        End If
        
    End With
    
Next

'MsgBox "移动处理完毕！"

End Sub



Sub 幻灯片设置标题()
'适用于幻灯片有标题文本，但不是标准标题布局的
'先设置幻灯片布局为标题幻灯片，此宏实现:查找最上面的文本，并设置为幻灯片标题
sn = Application.ActivePresentation.Slides.Count
Debug.Print sn

For i = 1 To sn

    Set myDocument = Application.ActivePresentation.Slides(i)
    rn = myDocument.Shapes.Count
    Debug.Print rn
    
    ActivePresentation.Slides(i).Select
       
    For Each s In ActivePresentation.Slides(i).Shapes
        'If s.Top > 17 And s.Top < 25 And s.Height < 20 Then
        '如果形状的top位置在指定范围、类型为17（文本框）
        If s.Top > -2 And s.Top < 25 And s.Type = 17 Then
            s.Select
            ' 取文本框第1段内容（去除段落标记）
            a = Replace(ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Paragraphs(1).Text, vbCr, "")
            
            If Not IsEmpty(a) Then
                Debug.Print a
                ActivePresentation.Slides(i).Shapes.Title.TextFrame2.TextRange.Text = a
                Else
                ActivePresentation.Slides(i).Shapes.Title.TextFrame2.TextRange.Text = ""
            End If
        End If
    Next

Next
 
MsgBox "处理完毕！"

End Sub



Sub 组合特定幻灯片中的shapes()

    Set myDocument = Application.ActivePresentation.Slides(6)
    rn = myDocument.Shapes.Count
    Debug.Print rn
    
    ActivePresentation.Slides(6).Shapes.SelectAll
    ActiveWindow.Selection.ShapeRange.Group

End Sub


Sub 组合每张幻灯片中的shapes()

sn = Application.ActivePresentation.Slides.Count
Debug.Print sn

For i = 1 To sn

    rn = Application.ActivePresentation.Slides(i).Shapes.Count
    Debug.Print rn
    ActivePresentation.Slides(i).Select
    ActivePresentation.Slides(i).Shapes.SelectAll
    ActiveWindow.Selection.ShapeRange.Group
    
Next

End Sub



Sub 奇数页移动上面的shapes()

    Set myDocument = Application.ActivePresentation.Slides(1)
    rn = myDocument.Shapes.Count
    Debug.Print rn
    
     ActivePresentation.Slides(1).Shapes.SelectAll
    
With ActiveWindow.Selection.ShapeRange
    
    .IncrementLeft -42.12
    .IncrementTop -94.2
    
    End With
    
End Sub



Sub 偶数页移动下面的shapes()

    Set myDocument = Application.ActivePresentation.Slides(1)
    rn = myDocument.Shapes.Count
    Debug.Print rn
    
     ActivePresentation.Slides(1).Shapes.SelectAll
    
With ActiveWindow.Selection.ShapeRange
    
    .IncrementLeft -42
    .IncrementTop -460
    
    End With
    
End Sub




Sub 删除多余形状()

sn = Application.ActivePresentation.Slides.Count
Debug.Print sn

For i = 1 To sn

    Set myDocument = Application.ActivePresentation.Slides(i)
    rn = myDocument.Shapes.Count
    Debug.Print rn
    
    ActivePresentation.Slides(i).Select
       
        ' 奇数页
        If i Mod 2 = 1 Then

            For Each s In ActivePresentation.Slides(i).Shapes
            If s.Top > 300 Then
                s.Delete
            End If
            Next
            
        Else
        '偶数页
            For Each s In ActivePresentation.Slides(i).Shapes
            If s.Top < 50 Then
                s.Delete
            End If
            Next
        End If

Next

MsgBox "处理完毕！"

End Sub



Sub 删除原标题文本()

sn = Application.ActivePresentation.Slides.Count
Debug.Print sn

For i = 1 To sn

    Set myDocument = Application.ActivePresentation.Slides(i)
    rn = myDocument.Shapes.Count
    Debug.Print rn
    
    ActivePresentation.Slides(i).Select
       
    For Each s In ActivePresentation.Slides(i).Shapes
        If s.Top > 5 And s.Top < 15 And s.Height < 25 Then
            s.Delete
        End If
    Next

Next
 
MsgBox "处理完毕！"

End Sub



Sub 删除ppt中特定大小的形状()

Dim currentSlide As Slide
Dim shp As Shape
 
For Each currentSlide In ActivePresentation.Slides
 For Each shp In currentSlide.Shapes

If shp.Width > 16 And shp.Width < 22 Then
shp.Delete

End If
 
 Next shp
Next currentSlide

End Sub







