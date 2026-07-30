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

' Инициализация: перехват событий + горячая клавиша + панель кнопок
Public Sub ИнициализироватьSmartCells()
    On Error Resume Next

    ' 1. Перехват событий приложения
    If gAppEvents Is Nothing Then Set gAppEvents = New AppEvents
    Set gAppEvents.App = Word.Application

    ' 2. Горячая клавиша Ctrl+Shift+U -> ОбновитьДанные
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyU), _
                    KeyCategory:=wdKeyCategoryCommand, _
                    Command:="ОбновитьДанные"

    ' 3. Панель кнопок SmartCells — появляется сама на вкладке «Надстройки»
    СоздатьПанельSmartCells
End Sub

' Создать панель с кнопками команд (вкладка «Надстройки» ленты Word).
' Так пользователю не нужно вручную настраивать панель быстрого доступа —
' достаточно нажать кнопку. Панель временная, пересоздаётся при старте Word.
Private Sub СоздатьПанельSmartCells()
    On Error Resume Next
    Dim bar As CommandBar
    Dim btn As CommandBarButton

    ' Удалить прежнюю панель, чтобы не плодить дубликаты
    Application.CommandBars("SmartCells").Delete

    Set bar = Application.CommandBars.Add(Name:="SmartCells", _
              Position:=msoBarTop, Temporary:=True)

    ' Кнопка «Обновить данные»
    Set btn = bar.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Обновить данные"
    btn.OnAction = "ОбновитьДанные"
    btn.Style = msoButtonCaption

    ' Кнопка «Настройки»
    Set btn = bar.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Настройки"
    btn.OnAction = "НастройкиSmartCells"
    btn.Style = msoButtonCaption
    btn.BeginGroup = True

    ' Кнопка «Убрать связь»
    Set btn = bar.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Убрать связь"
    btn.OnAction = "УбратьСвязьSmartCells"
    btn.Style = msoButtonCaption

    bar.Visible = True
End Sub
