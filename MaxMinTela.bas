Attribute VB_Name = "MaxMinTela"
Option Explicit

#If VBA7 Then
'// Condição aplicada a versão Office 64 bytes
Private Declare PtrSafe Function ExibirÍcone Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Function IniciaJanela Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function MoveJanela Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

#Else
'// Condição aplicada a versão Office 32 bytes
Private Declare Function ExibirÍcone Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function IniciaJanela Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function MoveJanela Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

#End If

'// Declaração de constantes para estilos de alteração e personalização dos Forms
Private Const FOCO_ICONE = &H80
Private Const ICONE = 0&
Private Const GRANDE_ICONE = 1&
Private Const ESTILO_PROLONGADO = (-20)
Private Const ESTILO_ATUAL As Long = (-16)
Private Const WS_CAPTION = &HC00000
Private Const WS_BARRA_TAREFAS = &H40000
Private Const WS_MENU As Long = &H80000
Private Const WS_CX_MINIMIZAR As Long = &H20000
Private Const WS_CX_MAXIMIZAR As Long = &H10000
Private Const WS_POPUP As Long = &H80000000
Private Const SW_EXIBIR_NORMAL = 1
Private Const SW_EXIBIR_MINIMIZADO = 2
Private Const SW_EXIBIR_MAXIMIZADO = 3
Dim Form_Personalizado As Long
Dim ESTILO As Long
Dim hIcone As Long


Sub Janelas()

Form_Personalizado = FindWindowA(vbNullString, UserForm1.Caption)

ESTILO = IniciaJanela(Form_Personalizado, ESTILO_ATUAL)

ESTILO = ESTILO Or WS_MENU
ESTILO = ESTILO Or WS_CX_MINIMIZAR
ESTILO = ESTILO Or WS_CX_MAXIMIZAR
ESTILO = ESTILO Or WS_POPUP '
ESTILO = ESTILO Or WS_CAPTION

MoveJanela Form_Personalizado, ESTILO_ATUAL, (ESTILO)

ESTILO = IniciaJanela(Form_Personalizado, ESTILO_PROLONGADO)
ESTILO = ESTILO Or WS_BARRA_TAREFAS

MoveJanela Form_Personalizado, ESTILO_PROLONGADO, ESTILO

hIcone = UserForm1.imgIcon.Picture.Handle
Call ExibirÍcone(Form_Personalizado, FOCO_ICONE, ICONE, ByVal hIcone)

DrawMenuBar Form_Personalizado

SetFocus Form_Personalizado
ShowWindow Form_Personalizado, 3
End Sub
