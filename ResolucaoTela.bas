Attribute VB_Name = "ResolucaoTela"
Option Explicit
 

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
#End If

Const SM_CXSCREEN = 0
Const SM_CYSCREEN = 1
 
Sub IndetificaResolucaoDaTela()
     
    Dim X  As Long
    Dim Y  As Long

    X = GetSystemMetrics(SM_CXSCREEN)
    Y = GetSystemMetrics(SM_CYSCREEN)
    
    Debug.Print X
    Debug.Print Y

End Sub

