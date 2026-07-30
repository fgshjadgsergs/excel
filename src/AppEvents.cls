VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =====================================================================
'  AppEvents Ч класс перехвата событий приложени€ Word.
'  —обытие DocumentOpen: если в открытом документе есть хот€ бы один
'  Content Control с тегом-префиксом "XL:" Ч выполнить автообновление.
'  Ёкземпл€р создаЄтс€ в AutoExec (модуль AutoExecInit).
' =====================================================================
Option Explicit

Public WithEvents App As Word.Application

' јвтообновление при открытии документа
Private Sub App_DocumentOpen(ByVal Doc As Document)
    On Error Resume Next
    Dim cc As ContentControl
    Dim hasXL As Boolean

    hasXL = False
    For Each cc In Doc.ContentControls
        If Left$(cc.Tag, 3) = "XL:" Then
            hasXL = True
            Exit For
        End If
    Next cc

    ' ≈сть метки-CC Ч запускаем главную команду обновлени€
    If hasXL Then ќбновитьƒанные
End Sub
