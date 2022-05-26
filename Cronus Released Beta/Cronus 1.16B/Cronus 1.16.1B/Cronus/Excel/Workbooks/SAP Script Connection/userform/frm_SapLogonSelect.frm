VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_SapLogonSelect 
   Caption         =   " Selecionar Mandante Sap"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5910
   OleObjectBlob   =   "frm_SapLogonSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_SapLogonSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Autor: Fabio Mitsueda
'Contato: mitsueda.fabio@gmail.com
'Data Criação: 12/08/2018
Option Explicit

Private Sub CmdConectar_Click()
    idSap = Me.cbo_Mandante.ListIndex + 1
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
