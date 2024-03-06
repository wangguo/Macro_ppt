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
    
    .IncrementLeft -42
    .IncrementTop -95
    
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


Sub 复制每页幻灯片()

sn = Application.ActivePresentation.Slides.Count
Debug.Print sn

For i = 1 To sn * 2 Step 2

    Application.ActivePresentation.Slides(i).Duplicate

Next

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
            .IncrementLeft -42
            .IncrementTop -95
        Else
        '偶数页
            .IncrementLeft -42
            .IncrementTop -460
        End If
        
    End With
    
Next

End Sub


Sub 删除形状()

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

End Sub


Sub 形状位置()

t = ActiveWindow.Selection.ShapeRange.Top
Debug.Print t

End Sub
