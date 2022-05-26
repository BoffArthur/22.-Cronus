VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Instalação do programa - CRONUS"
   ClientHeight    =   1710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7095
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
'
' Salvar
'

If Dir(Diretorio) = "saplogon.exe" Then
    msg = MsgBox("Instalação bem sucedida", vbOKOnly, "Sap Conect -Cronus")
    UserForm1.Hide
Else
    msg = MsgBox("Programa selecionado não identificado,Cronus não instalado com sucesso.", vbOKOnly, "Sap Conect - Cronus")
End If

End Sub

Private Sub CommandButton2_Click()
    Dim diaFolder As FileDialog
    Dim selected As Boolean

    ' Open the file dialog
    Set diaFolder = Application.FileDialog(msoFileDialogFilePicker) 'msoFileDialogFolderPicker
    diaFolder.AllowMultiSelect = False
    selected = diaFolder.Show

    If selected Then
        UserForm1.TextBox1.Value = diaFolder.SelectedItems(1) 'mostra ao usuário qual o diretorio escolhido
    End If

    Diretorio = UserForm1.TextBox1.Value 'salva o nome do diretorio escolhido

    Set diaFolder = Nothing
End Sub

Private Sub CommandButton3_Click()

msg = MsgBox("Para que o CAS funcione corretamente é necessária a informação do local de instalação do SAP." & vbCrLf & _
              "Programa procurado é 'SAPLogon.exe'. Diretorio normalmente encontrado:" & vbCrLf & vbCrLf & _
              "C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe" & vbCrLf & vbCrLf & vbCrLf & _
              "Dúvida contatar o administrador do CAS.", vbOKOnly, "Sap Conect - Cronus")
End Sub

Private Sub CommandButton4_Click()
'
' Cancelar
'

UserForm1.Hide

End Sub

Private Sub Label1_Click()

End Sub

Private Sub TextBox1_Change()

Diretorio = TextBox1.Text

End Sub

Private Sub UserForm_Click()

End Sub
