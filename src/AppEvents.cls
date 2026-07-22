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
'  AppEvents — класс перехвата событий приложения Word.
'  Событие DocumentOpen: если в открытом документе есть хотя бы один
'  Content Control с тегом-префиксом "XL:" — выполнить автообновление.
'  Экземпляр создаётся в AutoExec (модуль AutoExecInit).
' =====================================================================
Option Explicit

Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1

' Автообновление при открытии документа
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

    ' Есть метки-CC — запускаем главную команду обновления
    If hasXL Then ОбновитьДанные
End Sub
