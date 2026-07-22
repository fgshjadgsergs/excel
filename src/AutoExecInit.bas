Attribute VB_Name = "AutoExecInit"
' =====================================================================
'  AutoExecInit — инициализация надстройки при старте Word.
'  Решение живёт в глобальном шаблоне (папка STARTUP), поэтому
'  AutoExec выполняется автоматически при запуске Word:
'    - создаётся экземпляр AppEvents (перехват DocumentOpen);
'    - назначается сочетание Ctrl+Shift+U на команду «ОбновитьДанные».
' =====================================================================
Option Explicit

' Экземпляр класса событий должен жить всю сессию Word,
' поэтому храним его в переменной уровня модуля.
Private gAppEvents As AppEvents

' Автозапуск при старте Word (для шаблона из STARTUP)
Public Sub AutoExec()
    ИнициализироватьSmartCells
End Sub

' Инициализация: перехват событий + горячая клавиша
Public Sub ИнициализироватьSmartCells()
    On Error Resume Next

    ' 1. Перехват событий приложения
    If gAppEvents Is Nothing Then Set gAppEvents = New AppEvents
    Set gAppEvents.App = Word.Application

    ' 2. Горячая клавиша Ctrl+Shift+U -> ОбновитьДанные (опционально)
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyU), _
                    KeyCategory:=wdKeyCategoryCommand, _
                    Command:="ОбновитьДанные"
End Sub
