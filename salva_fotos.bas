Attribute VB_Name = "Módulo1"
Sub Macro1()
'
' Macro1 Macro
'

'
    For Each oShape In ActiveSheet.Shapes
    strImageName = ActiveSheet.Cells(oShape.TopLeftCell.Row, 1).Value
    oShape.Select
    'Picture format initialization
    Selection.ShapeRange.PictureFormat.Contrast = 0.5:
    Selection.ShapeRange.PictureFormat.Brightness = 0.5:
    Selection.ShapeRange.PictureFormat.ColorType = msoPictureAutomatic:
    Selection.ShapeRange.PictureFormat.TransparentBackground = msoFalse:
    Selection.ShapeRange.Fill.Visible = msoFalse
    Selection.ShapeRange.Line.Visible = msoFalse
    Selection.ShapeRange.Rotation = 0#
    Selection.ShapeRange.PictureFormat.CropLeft = 0#
    Selection.ShapeRange.PictureFormat.CropRight = 0#
    Selection.ShapeRange.PictureFormat.CropTop = 0#
    Selection.ShapeRange.PictureFormat.CropBottom = 0#
    Selection.ShapeRange.ScaleHeight 1#, msoTrue, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleWidth 1#, msoTrue, msoScaleFromTopLeft
    '/Picture format initialization
    
    Application.Selection.CopyPicture
    Set oDia = ActiveSheet.ChartObjects.Add(0, 0, oShape.Width, oShape.Height)
    Set oChartArea = oDia.Chart
    oDia.Activate
    
    With oChartArea
        .ChartArea.Select
        .Paste
        .Export ("c:\fotos\" & strImageName & ".jpg")
    End With
    
    oDia.Delete 'oChartArea.Delete
Next
End Sub



Sub SavaFotos()
Attribute SavaFotos.VB_ProcData.VB_Invoke_Func = "Y\n14"
'
' SavaFotos Macro
'
' Atalho do teclado: Ctrl+Shift+Y
'
Dim shp As Shape
Dim ws As Worksheet
Dim initialWs As String
Dim pathFiles As String

initialWs = ActiveSheet.Name ' Salve o nome da worksheet inicial
'pathFiles = Application.InputBox("Quer salvar onde?")

Dim sFolder As String
' Open the select folder prompt
With Application.FileDialog(msoFileDialogFolderPicker)
    If .Show = -1 Then ' if OK is pressed
        sFolder = .SelectedItems(1)
    End If
End With

If sFolder <> "" Then ' if a file was chosen
    For Each ws In Sheets 'Navega entre as worksheets.
        ws.Activate
        
        For Each shp In ws.Shapes 'Navega entre as shapes da worksheet atual
    
            If shp.Type = msoPicture Then 'É imagem?
                shp.Select
                
                Application.Selection.CopyPicture 'Copia a imagem
                
                Set oDia = ActiveSheet.ChartObjects.Add(0, 0, shp.Width, shp.Height) ' Adiciona um ChartObject
                Set oChartArea = oDia.Chart
                
                oDia.Activate
                
                Application.Selection.Height = 600 'Muda a altura do chart
                Application.Selection.Width = 800 'Muda a largura do chart
                
                With oChartArea ' Cola a imagem no chart
                    .ChartArea.Select
                    .Paste
                End With
                
                Selection.ShapeRange.Height = 600 'Muda a altura da imagem dentro do chart
                Selection.ShapeRange.Width = 800 'Muda a lrgura da imagem dentro do chart
                
                With oChartArea 'Salva a imagem como jpg
                    .Export (sFolder & "/" & ws.Name & "-" & shp.Name & ".jpg")
                End With
                
                oDia.Delete 'Deleta o chart
        End If
    
        Next shp
    Next ws
    
    Worksheets(initialWs).Activate 'Volta pra worksheet inicial
    
    MsgBox "Terminei! ;)"

End If

End Sub

