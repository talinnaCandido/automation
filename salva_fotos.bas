
Sub SavaFotos()
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
    
            If shp.Type = msoPicture Then 'imagem?
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

