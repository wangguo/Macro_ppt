Sub 裁剪图并调整位置片()
    Dim mySlide As Slide
    Dim myShape As Shape, i_Temp As Integer
    On Error Resume Next
    
    
    For Each mySlide In ActivePresentation.Slides

        For Each myShape In mySlide.Shapes
            If myShape.Type = msoPicture Then
            
                With myShape
                ' 重置图片尺,msoTrue为相对原始尺寸缩放
               .ScaleHeight 1, msoTrue
               .ScaleWidth 1, msoTrue

                 .PictureFormat.CropLeft = 57.06
               .PictureFormat.CropTop = 122.87
                 '.PictureFormat.CropRight = 10
                '.PictureFormat.CropBottom = 67
                
    
                .Left = 0
                .Top = 0
                ' .Height = 810
                ' .Width = 1440
                End With
            End If
        Next
    Next
End Sub


Sub 按奇偶页裁剪图片()
    Dim mySlide As Slide
    Dim myShape As Shape, i_Temp As Integer
    On Error Resume Next
    
    For Each mySlide In ActivePresentation.Slides

        For Each myShape In mySlide.Shapes
            If myShape.Type = msoPicture Then
            
                With myShape
                ' 重置图片尺,msoTrue为相对原始尺寸缩放
               .ScaleHeight 1, msoTrue
               .ScaleWidth 1, msoTrue
               
               i = mySlide.SlideNumber
               Debug.Print i
               
        
                ' 奇数页
                If i Mod 2 = 1 Then
                    .PictureFormat.CropLeft = 57.06
                    .PictureFormat.CropTop = 122.87
                     '.PictureFormat.CropRight = 10
                     '.PictureFormat.CropBottom = 67
            Else
            '偶数页
                    .PictureFormat.CropLeft = 57.06
                   .PictureFormat.CropTop = 52.5
                     '.PictureFormat.CropRight = 10
                    '.PictureFormat.CropBottom = 67
            End If
            
                .Left = 0
                .Top = 0
                ' .Height = 810
                ' .Width = 1440
                End With
            End If
        Next
    Next
End Sub




Sub 指定ppt裁剪图片()
    ' 定义变量
    Dim pptFilePath As String
    Dim presentation As presentation
      
    ' 设置 pptx 文件路径
    pptFilePath = "/Users/xxxx.pptx"
      
    ' 打开 pptx 文件
    Set presentation = Presentations.Open(pptFilePath)
    
    ' 调整页面尺寸(72前面的数字单位为英寸)
    
   With Application.ActivePresentation.PageSetup
         .SlideWidth = 20 * 72
        .SlideHeight = 11.2519685 * 72
  End With
      
    处理图片
      
        ' 保存并关闭 pptx 文件
        presentation.Save
        presentation.Close
        
End Sub



Sub 批量调整PPT()
    ' 定义变量
    Dim folderPath As String
    Dim presentation As presentation
      
    ' 设置文件夹路径
    folderPath = "/Users/xxx/"
      
    ' 打开文件夹内的所有 pptx 文件
    Dim fileName As String
    fileName = Dir(folderPath & "*.pptx")
    Do While fileName <> ""
        ' 打开 pptx 文件
        Set presentation = Presentations.Open(folderPath & fileName)
           
        ' 调整页面尺寸(72前面的数字单位为英寸)
        With Application.ActivePresentation.PageSetup
             .SlideWidth = 20 * 72
            .SlideHeight = 11.2519685 * 72
        End With
  
  
        指定ppt裁剪图片
          
        ' 保存并关闭 pptx 文件
        presentation.Save
        presentation.Close
          
        ' 获取下一个文件名
        fileName = Dir()
    Loop
End Sub






Sub 合并PPT()

    Dim objPPT As PowerPoint.Application
    Dim objPresentation As PowerPoint.presentation
    Dim strFile As String, strFolder As String
    strFolder = "/Users/xxx" '设置需要合并的PPT文件所在文件夹路径
    Set objPPT = New PowerPoint.Application
    Set objPresentation = objPPT.Presentations.Add(msoTrue)
    strFile = Dir(strFolder & "/*.pptx") '遍历文件夹中的所有PPT文件

    Do While strFile <> ""
        objPresentation.Slides.InsertFromFile strFolder & "/" & strFile, objPresentation.Slides.Count
        strFile = Dir()
    Loop
    objPresentation.SaveAs "/Users/wangguo/python/PPT/PPT标准化/hb/hb.pptx", ppSaveAsOpenXMLPresentation '保存合并后的文件

End Sub



