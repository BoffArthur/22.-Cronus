Attribute VB_Name = "UDM"
' UDM = User Defined Macro

Sub SelectArchive()
    Dim diaFolder As FileDialog
    Dim selected As Boolean

    ' Open the file dialog
    Set diaFolder = Application.FileDialog(msoFileDialogFilePicker)
    diaFolder.AllowMultiSelect = False
    selected = diaFolder.Show
    
    If selected Then                              'se for selecionado o arquivo
        ThisWorkbook.Sheets("Listas de Dados").Range("A7").Value = diaFolder.SelectedItems(1) 'Salvar o diretorio do arquivo na aba lista de dados.
        Call AtualizarPQ
    End If
    
    Set diaFolder = Nothing
End Sub
Sub SelectThisArchive()
    '
    ' DEfinir o arquivo atual como origem de consulta do PowerQuery
    '
    Dim FilePath As String

    Application.ActiveWorkbook.Save
    
    FilePath = Application.ActiveWorkbook.path    'salva o local do arquivo ativo
    
    If FilePath = "" Then                         'se o arquivo nao estiver salvo em local algum, retornar a msg de erro ao usuário
        msgem = MsgBox("Local de salvamento deste arquivo não encontrado. As seguintes soluções são possíveis:" & vbCrLf & _
                "" & vbCrLf & _
                "- Salvar a pasta de trabalho." & vbCrLf & _
                "- Alterar a pasta raiz da pasta de trabalho para um local válido." & vbCrLf & _
                "" & vbCrLf & _
                "Para mais informações contatar o Administrador do sistema.", vbCritical, "Cronus")
        ThisWorkbook.Close
        Exit Sub
    Else                                          ' Caso esteja salvo então continuar a consulta
        FilePath = Application.ActiveWorkbook.FullName
        ThisWorkbook.Sheets("Listas de Dados").Range("A7").Value = FilePath 'define a origem da consulta do PowerQuery como o local do arquivo ativo
        Call AtualizarPQ
    End If
    
End Sub
Sub AtualizarPQ()
    '
    ' Atualizar Power Query da consulta
    '
    On Error GoTo ErrorHandlerMSG
    Dim RowValue As String
    Dim I      As String
    
    Application.ScreenUpdating = False            'deativa a atualização de tela
    
    ThisWorkbook.Worksheets("Consulta").ListObjects("Estrutura_P01").QueryTable.Refresh BackgroundQuery:=False 'Atualizar Power Query com a consulta de estrutura
    ThisWorkbook.Worksheets("Consulta").ListObjects("Estrutura_P01").Range.AutoFilter Field:=2, Criteria1:="NEW" 'Filtrar a consulta somente com novos cadastros
    
    Application.ScreenUpdating = True             'ativa a atualização de tela
    UserForm3.Show
    Exit Sub
    
ErrorHandlerMSG:
    MsgBox1 = MsgBox("Ocorreu um erro ao fazer a consulta com o Power Query do arquivo desejado. " & vbCrLf & _
              "Os seguintes erros são possiveis:" & vbCrLf & _
              "" & vbCrLf & _
              "- A pasta raiz está fora do modelo padrão aceitável. Deve-se utilizar o arquivo padrão do Cronus para tal." & vbCrLf & _
              "- As permissões de privacidade das fórmulas do Power Query estão mais restritas do que o necessário para o Cronus. Deve-se alterar para um nível menor dentro do Power Query." & vbCrLf & _
              "" & vbCrLf & _
              "Para mais informações contatar o Administrador do sistema.", vbCritical, "Cronus")
    
End Sub

Sub LoadFromMM01()
    '
    ' Transferir os códigos da MM01 para aba atual
    '

    Dim Aba    As String                          'define a variavel para a aba atual do usuário

    Aba = ThisWorkbook.ActiveSheet.Name           'salva o nome da aba atual do usuário
    If Aba = "Consulta" Or Aba = "Listas de Dados" Then Exit Sub 'se a aba atual não for o nome de uma transação, sair da macro

    I = ThisWorkbook.Sheets("MM01").Range("b1048576").End(xlUp).Row 'verifica quantas linhas tem preenchidas na aba MM01

    Do Until I = 4                                'copia todos os códigos
    Sheets(Aba).Range("B" & I).Value = ThisWorkbook.Sheets("MM01").Range("b" & I).Value
    I = I - 1
Loop

End Sub

Public Sub SalvarMSG(Função As String, I, Trsc)
    '
    ' Salvar a menssagem final do cadastro
    ' Função = Salvar ou Repetir
    '
    If Função = "Salvar" Then                     'salva a menssagem final somente na linha atual alterada

    ThisWorkbook.Sheets(Trsc).Cells(I, WorksheetFunction.Match("MsgHandler", ThisWorkbook.Sheets(Trsc).Range("3:3"), 0) + 0).Value = _
                                       Session.findById("wnd[0]/sbar").Text & " - em " & Format(Now, "DD/MM/YY HH:MM:SS")
    
ElseIf Função = "Repetir" Then                    'salva a menssagem final na linha atual alterada e repete para todas as linhas acima
    I = I - 1
    Do Until I = 4                                'salvar a msg em todos os itens atualizados
        'Salvar a menssagem final ao sair do cadastro
        ThisWorkbook.Sheets(Trsc).Cells(I, WorksheetFunction.Match("MsgHandler", ThisWorkbook.Sheets(Trsc).Range("3:3"), 0) + 0).Value = _
                                           Session.findById("wnd[0]/sbar").Text & " - em " & Format(Now, "DD/MM/YY HH:MM:SS")
        I = I - 1
    Loop
End If

End Sub

Sub ShowUserForm()
    '
    ' A macro principal deve ficar ligada ao aparecimento da Userform2, já que ela vai mostrar a %done da macro
    '
    UserForm2.Show
End Sub

Sub UpdateProgressBar(PctDone As Single, CycleTime As String)
    '
    ' Macro para atualizar a %done do scrip.
    '


    'lida com a opção do tempo, fazendo a conversão em segundos, minutos e horas
    If CycleTime < 60 Then
        CycleTime = Round(CycleTime, "0") & " Seg."
    ElseIf CycleTime > 60 Then
        CycleTime = Round((CycleTime / 60), "2") & " Min."
    ElseIf CycleTime > 3600 Then
        CycleTime = Round((CycleTime / 60) / 60, "2") & " hrs"
    End If
    
    With UserForm2
        ' Update the Caption property of the Frame control.
        .FrameProgress.Caption = Format(PctDone, "0%") 'Atualiza o nome do Frame para a %done.
        
        ' Widen the Label control.
        .LabelProgress.Width = PctDone * (.FrameProgress.Width - 10) 'atualiza a largura da Label com base na %done.
        
        ' Atualiza o tempo de clico restante total em minutos
        .TimeProgress.Caption = CycleTime
    End With
    
    ' The DoEvents allows the UserForm to update. A planilha inteira vai se atualizar, não somente a UserForm
    DoEvents
End Sub

Sub GerarRelatorio()
    '
    ' Gerar um relatorio final de todas as criações/alterações feitas com o scrip
    '
    Dim WS_Count As Integer
    Dim I      As Integer
    Dim wb     As Workbook                        ' Definir uma variavel para o WORKBOOK
    Dim LocalNome As String                       ' Definir o local com o nome
    Dim diaFolder As FileDialog
    Dim selected As Boolean
    Dim ColValue As String
    Dim ColCnt As Integer

    Application.ScreenUpdating = False
    
    ' Set WS_Count equal to the number of worksheets in the active workbook
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    ' Begin the loop.
    
    
    ' Open the file dialog
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    selected = diaFolder.Show
    
    If selected Then                              'se for selecionado o arquivo
        Sheets("Listas de Dados").Range("A10").Value = diaFolder.SelectedItems(1) 'Salvar o diretorio do arquivo na aba lista de dados.
        LocalNome = Sheets("Listas de Dados").Range("A10").Value 'define a variavel LocalNome como o local que deve ser salva a planilha
        Set wb = Workbooks.Add                    ' Cria uma nova planilha (WORKBOOK)
        Sheets(1).Name = "Relatorio Script"
        ActiveWindow.DisplayGridlines = False
        
        For I = 1 To WS_Count
            
            ThisWorkbook.Activate                 ' Ativa este workbook
            ColCnt = 1
            Do Until ThisWorkbook.Sheets(I).Cells(3, ColCnt).Value = "MsgHandler"
                ColCnt = ColCnt + 1
                If ColCnt > 1000 Then GoTo Salto1
            Loop
            
            ThisWorkbook.Sheets(I).Columns(ColCnt).Copy ' Copiar a planilh LIT para o novo arquivo criado
            wb.Activate                           ' Ativa a planilha criada
            Application.DisplayAlerts = False     ' Desliga o alerta para não perguntar se quer deletar a planilha
            
            
            'procurar a primeira coluna em branco
            ColCnt = 1
            Do Until Sheets("Relatorio Script").Cells(1, ColCnt).Value = ""
                ColCnt = ColCnt + 1
            Loop
            
            'colar os dados copiados na primeira coluna livre
            Sheets("Relatorio Script").Cells(1, ColCnt).Select
            ActiveSheet.Paste
            Sheets("Relatorio Script").Cells(1, ColCnt).Value = ThisWorkbook.Sheets(I).Name
            
            Application.DisplayAlerts = True      ' Liga novamente os alertas
            
Salto1:
        Next I
        On Error Resume Next
        LocalNome = LocalNome & "\" & Format(Date, "yyyy-mm-dd") & "_Relatório Script - " & _
                    Format(Now, "yymmdd hhmmss") & ".xlsx"
        wb.Sheets(1).Rows("2:3").Delete           'deletar as duas linhas com informações inuteis na planilha criada
        wb.SaveAs LocalNome                       ' Salva no local predeterminado com o nome já escolhido
        wb.Close                                  'fecha a planilha criada
    End If
    Application.ScreenUpdating = True
End Sub

Sub OpenThisWB()
    '
    '
    '

    ThisWorkbook.Activate
End Sub

