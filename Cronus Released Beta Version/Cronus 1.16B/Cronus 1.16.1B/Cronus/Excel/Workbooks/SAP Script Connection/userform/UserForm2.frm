VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Executando Scrip. Aguarde..."
   ClientHeight    =   855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5685
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
    If AutomaticRun <> vbYes Then Call ExecutarScript 'ao ativar a userform2, come�ar a executar o scrip
    If ActiveSheet.Index = 1 Then Call ExecutarScript
End Sub
