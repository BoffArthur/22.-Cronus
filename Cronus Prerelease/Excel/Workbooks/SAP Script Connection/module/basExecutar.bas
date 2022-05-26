Attribute VB_Name = "basExecutar"
Option Explicit
Public Diretorio As String
Public AutomaticRun As Variant

Public Sub ExecutarScript()
    Dim intRetVal   As Integer
    Dim vSap        As Variant
    Dim X           As Integer
    Dim Transação   As String

    'Verificando se existe SAP logado
    vSap = GetSapUserOpen()
    
    Transação = ThisWorkbook.ActiveSheet.Name
    Diretorio = UserForm1.TextBox1.Text
    
    
    If Diretorio = "" Then
        UserForm1.Show
        Exit Sub
    End If
    
    If AutomaticRun <> "" Then
        If AutomaticRun = vbYes And ActiveSheet.Index > 1 Then
            'Se tem janelas SAP na tela inicial ou seja logada e em nenhuma transação execute abaixo
            If vSap(1)(4) = True Then
                Set objConnection = vSap(1)(5)
                Set Session = vSap(1)(6)
                'Procurando a janela incial dentro das possiveis janelas abertas nesta conexão
                For Each Session In objConnection.Children
                    If Session.Busy = False Then
                        If Session.Info.Transaction = "SESSION_MANAGER" Then
                            Call Executar(Transação)
                            Exit For
                        End If
                    End If
                Next
                'Caso não tenha uma janela na tela inical criar uma janela para utilizar a transação
            Else
                Set objConnection = vSap(1)(5)
                Set Session = vSap(1)(6)
                'Criando uma janela nova, que vai estar na tela inicial
                Session.createsession
                'É necessario pausar um pouco o codigo para dar tempo da janela ser aberta
                Application.Wait Now + TimeValue("00:00:6")
                'Procurando a janela incial que acabamos de criar
                For Each Session In objConnection.Children
                    If Session.Busy = False Then
                        If Session.Info.Transaction = "SESSION_MANAGER" Then
                            Call Executar(Transação)
                            Exit For
                        End If
                    End If
                Next
            End If
            
            Set objSapGui = Nothing
            Set objApplication = Nothing
            Set objConnection = Nothing
            Set Session = Nothing
            AutomaticRun = ""
            Exit Sub
        End If
    End If
    'vSap é uma variante que na linha anterior recebeu valores da função GetSapUserOpen, quando não existe nenhuma instancia Sap aberta
    'Essa função retorna Empty, ou seja nada, na linha abaixo testo se a função não retornou nada é porque existe conexões abertas.
    If Not IsEmpty(vSap) Then
        'A função UBound retorna o ultimo campo de uma matriz, caso a matriz seja maior que 1 isso significa que existem mais de uma instancia de Sap aberta
        If UBound(vSap) > 1 Then
            'Caso existe mais de um mandante aberto será aberto a opção de selecionar um mandante ou efetuar um novo login
            intRetVal = MsgBox("Existem mais de um mandante aberto!" & vbCrLf & vbCrLf & _
                        "Para selecionar um mandante clique em Sim." & vbCrLf & _
                        "Para conectar a um novo mandante clique em Não." & vbCrLf & _
                        "Ou Cancelar para sair.", vbYesNoCancel + vbInformation, "Sap Conect - Cronus")
            
            'Caso deseje selecionar um mandante abrir form de seleção
            If intRetVal = vbYes Then
                idSap = 0
                Load frm_SapLogonSelect
                With frm_SapLogonSelect.cbo_Mandante
                    For X = LBound(vSap) To UBound(vSap)
                        .AddItem "Id: " & vSap(X)(2) & " | Usuario: " & vSap(X)(3)
                    Next
                End With
                frm_SapLogonSelect.Show
                If idSap > 0 Then
                    'Se tem janelas SAP na tela inicial ou seja logada e em nenhuma transação execute abaixo
                    If vSap(idSap)(4) = True Then
                        Set objConnection = vSap(idSap)(5)
                        Set Session = vSap(idSap)(6)
                        'Procurando a janela incial dentro das possiveis janelas abertas nesta conexão
                        For Each Session In objConnection.Children
                            If Session.Busy = False Then
                                If Session.Info.Transaction = "SESSION_MANAGER" Then
                                    Call Executar(Transação)
                                    Exit For
                                End If
                            End If
                        Next
                        'Caso não tenha uma janela na tela inical criar uma janela para utilizar a transação
                    Else
                        Set objConnection = vSap(idSap)(5)
                        Set Session = vSap(idSap)(6)
                        'Criando uma janela nova, que vai estar na tela inicial
                        Session.createsession
                        'É necessario pausar um pouco o codigo para dar tempo da janela ser aberta
                        Application.Wait Now + TimeValue("00:00:6")
                        'Procurando a janela incial que acabamos de criar
                        For Each Session In objConnection.Children
                            If Session.Busy = False Then
                                If Session.Info.Transaction = "SESSION_MANAGER" Then
                                    Call Executar(Transação)
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                End If
                'Caso clique em não abrir formulario de nova conexão
            ElseIf intRetVal = vbNo Then
                'Load frm_SapLogon
                'frm_SapLogon.Show
                UserForm2.Hide
            ElseIf intRetVal = vbCancel Then
                UserForm2.Hide
            End If
        Else
            'Perguntando ao usuario se ele deseja utilizar a conexão aberta
            intRetVal = MsgBox("Deseja efetuar essa operação com os dados abaixo?" & vbCrLf & vbCrLf & _
                        "Sistema: " & vSap(1)(2) & vbCrLf & _
                        "Usuário: " & vSap(1)(3), vbYesNoCancel + vbInformation, "Sap Conect - Cronus")
            
            'Caso o usuario deseje utiliza a conexão indicada intRetVal será igual a sim e sera executado o trecho abaixo
            If intRetVal = vbYes Then
                'Se tem janelas SAP na tela inicial ou seja logada e em nenhuma transação execute abaixo
                If vSap(1)(4) = True Then
                    Set objConnection = vSap(1)(5)
                    Set Session = vSap(1)(6)
                    'Procurando a janela incial dentro das possiveis janelas abertas nesta conexão
                    For Each Session In objConnection.Children
                        If Session.Busy = False Then
                            If Session.Info.Transaction = "SESSION_MANAGER" Then
                                Call Executar(Transação)
                                Exit For
                            End If
                        End If
                    Next
                    'Caso não tenha uma janela na tela inical criar uma janela para utilizar a transação
                Else
                    Set objConnection = vSap(1)(5)
                    Set Session = vSap(1)(6)
                    'Criando uma janela nova, que vai estar na tela inicial
                    Session.createsession
                    'É necessario pausar um pouco o codigo para dar tempo da janela ser aberta
                    Application.Wait Now + TimeValue("00:00:6")
                    'Procurando a janela incial que acabamos de criar
                    For Each Session In objConnection.Children
                        If Session.Busy = False Then
                            If Session.Info.Transaction = "SESSION_MANAGER" Then
                                Call Executar(Transação)
                                Exit For
                            End If
                        End If
                    Next
                End If
                'Caso o usuario não deseje utiliza a conexão indicada faça abertura de um formulario de conexão
            ElseIf intRetVal = vbNo Then
                'Load frm_SapLogon
                'frm_SapLogon.Show
                UserForm2.Hide
            ElseIf intRetVal = vbCancel Then
                UserForm2.Hide
            End If
        End If
        
        'Descarregando todos os objetos da memoria
        For X = LBound(vSap) To UBound(vSap)
            Set vSap(X)(5) = Nothing
            Set vSap(X)(6) = Nothing
        Next
        
        'Caso não encontre nenhuma instancia Sap aberta solicitar login pelo userform
    Else
        'Load frm_SapLogon
        'frm_SapLogon.Show
    End If
    
    'Descarregando todos os objetos da memoria
    Set objSapGui = Nothing
    Set objApplication = Nothing
    Set objConnection = Nothing
    Set Session = Nothing
    
End Sub

