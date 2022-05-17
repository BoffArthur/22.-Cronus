VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_SapLogon 
   Caption         =   " Conectar e Executar Script"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5730
   OleObjectBlob   =   "frm_SapLogon.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_SapLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Autor: Fabio Mitsueda
'Contato: mitsueda.fabio@gmail.com
'Data Criação: 12/08/2018
Option Explicit

Private Sub CmdConectar_Click()
Dim strUser          As String   'Usuário informado no momento da conexão
Dim strPassword      As String   'Senha do úsuario informada no momento da conexão
Dim strAmbiente      As String   'Ambiente SAP indicado no momento da conexão

    'Consiste campos
    If Me.txt_Usuario = "" Or _
        Me.txt_Senha = "" Or _
        Me.txt_Ambiente = "" Then
        MsgBox "Informe todos os campos.", vbCritical, "Conexão com o SAP"
        Exit Sub
    End If
    
    'Carregando variaveis de ambiente, usuario e senha
    strUser = Me.txt_Usuario
    strPassword = Me.txt_Senha
    strAmbiente = Me.txt_Ambiente
   
    If ConectarSAP(strUser, strPassword, strAmbiente) Then
        'Chamando transação
        Call Executar
    End If
       
    On Error Resume Next
        AppActivate "Microsoft Excel"
    On Error GoTo 0
    
    'Descarregando userform
    Unload Me

End Sub

'Executar a rotina de conexeção ao teclar enter no campo senha
Private Sub txt_Senha_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CmdConectar_Click
    End If
End Sub

Private Sub UserForm_Click()

End Sub
