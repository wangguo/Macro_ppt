Sub 处理图片()
    Dim mySlide As slide
    Dim myShape As Shape, i_Temp As Integer
    On Error Resume Next
        
    For Each mySlide In ActivePresentation.Slides

        For Each myShape In mySlide.Shapes
            If myShape.Type = msoPicture Then
            
                With myShape
                ' 重置图片尺,msoTrue为相对原始尺寸缩放
               .ScaleHeight 1, msoTrue
               .ScaleWidth 1, msoTrue

                 '.PictureFormat.CropLeft = 10
               .PictureFormat.CropTop = 23.3
                 '.PictureFormat.CropRight = 10
                .PictureFormat.CropBottom = 67
                               
                .Left = 0
                .Top = 0
                ' .Height = 810
                ' .Width = 1440
                End With
            End If
        Next
    Next
End Sub

'************************************
Sub 调整指定ppt的页面和图片尺寸()
    ' 定义变量
    Dim pptFilePath As String
    Dim presentation As presentation
      
    ' 设置 pptx 文件路径
    pptFilePath = "/Users/xxx.pptx"
      
    ' 打开 pptx 文件
    Set presentation = Presentations.Open(pptFilePath)
    
    ' 调整页面尺寸(72前面的数字单位为英寸)
    
   With Application.ActivePresentation.PageSetup
         .slideWidth = 20 * 72
        .slideHeight = 11.2519685 * 72
  End With
      
    处理图片
      
        ' 保存并关闭 pptx 文件
        presentation.Save
        presentation.Close
        
End Sub


'************************************
Sub 批量调整PPT()
    ' 定义变量
    Dim folderPath As String
    Dim presentation As presentation
      
    ' 设置文件夹路径
    folderPath = "/Users/ppt文件夹/"
      
    ' 打开文件夹内的所有 pptx 文件
    Dim fileName As String
    fileName = Dir(folderPath & "*.pptx")
    Do While fileName <> ""
        ' 打开 pptx 文件
        Set presentation = Presentations.Open(folderPath & fileName)
           
        ' 调整页面尺寸(72前面的数字单位为英寸)
        With Application.ActivePresentation.PageSetup
             .slideWidth = 20 * 72
            .slideHeight = 11.2519685 * 72
        End With
  
  
        处理图片
          
        ' 保存并关闭 pptx 文件
        presentation.Save
        presentation.Close
          
        ' 获取下一个文件名
        fileName = Dir()
    Loop
End Sub


'************************************


