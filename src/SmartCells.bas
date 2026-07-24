Attribute VB_Name = "SmartCells"
' =====================================================================
'  SmartCells — подстановка значений ячеек Excel в текст документа Word
'  Основной модуль: команды, разбор меток, работа с Excel, сводка.
'  Механика: метки {адрес} превращаются в Content Control с тегом
'  "XL:адрес"; повторное обновление идёт по этим CC.
' =====================================================================
Option Explicit

' --- Раздел реестра и ключи настроек ---
Private Const REG_APP As String = "NODA_SmartCells"
Private Const REG_SECTION As String = "Settings"
Private Const KEY_PATH As String = "ExcelPath"
Private Const KEY_SHEET As String = "SheetName"

' --- Параметры Content Control ---
Private Const TAG_PREFIX As String = "XL:"     ' префикс тега метки-CC
Private Const CC_TITLE As String = "SmartCells"

' --- Цвета заливки (константы Word) ---
Private Const CLR_YELLOW As Long = 65535           ' wdColorYellow
Private Const CLR_AUTO As Long = -16777216         ' wdColorAutomatic (без заливки)

' =====================================================================
'  ГЛАВНАЯ КОМАНДА: обновить все значения в документе
' =====================================================================
Public Sub ОбновитьДанные()
    Dim sPath As String, sSheet As String
    Dim xlApp As Object, xlWb As Object, xlWs As Object
    Dim doc As Document
    Dim updated As Long
    Dim aborted As Boolean
    Dim problems As Collection
    Set problems = New Collection

    ' 1. Прочитать настройки, при отсутствии — вызвать диалог
    If Not ПрочитатьНастройки(sPath, sSheet) Then
        НастройкиSmartCells
        If Not ПрочитатьНастройки(sPath, sSheet) Then
            MsgBox "Настройки не заданы. Обновление отменено.", _
                   vbExclamation, "SmartCells"
            Exit Sub
        End If
    End If

    ' 2. Проверить наличие файла Excel — если нет, документ не трогаем вообще
    If Len(Dir(sPath)) = 0 Then
        MsgBox "Файл Excel не найден по пути:" & vbCrLf & vbCrLf & _
               sPath & vbCrLf & vbCrLf & _
               "Проверьте путь через команду «НастройкиSmartCells».", _
               vbExclamation, "SmartCells"
        Exit Sub
    End If

    Set doc = ActiveDocument
    Application.ScreenUpdating = False

    ' 3. Открыть Excel один раз на весь проход (late binding, ReadOnly, невидимо)
    On Error GoTo Fail
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    Set xlWb = xlApp.Workbooks.Open(FileName:=sPath, ReadOnly:=True, AddToMru:=False)

    ' Найти нужный лист; если листа нет — документ не меняем
    Set xlWs = НайтиЛист(xlWb, sSheet)
    If xlWs Is Nothing Then
        MsgBox "Лист «" & sSheet & "» не найден в книге:" & vbCrLf & _
               sPath & vbCrLf & vbCrLf & _
               "Проверьте имя листа через «НастройкиSmartCells».", _
               vbExclamation, "SmartCells"
        aborted = True
        GoTo CleanUp
    End If

    ' 4. Проход (а): текст {адрес} -> Content Control (проблемные метки подсветить)
    ПройтиПоТексту doc, problems

    ' 5. Проход (б): перечитать значения всех CC XL:* из Excel
    '    (лист по умолчанию — xlWs; метки {Лист!A5} читаются из своих листов)
    updated = ПройтиПоCC(doc, xlWb, xlWs, problems)

CleanUp:
    ' --- Гарантированное закрытие Excel и освобождение объектов ---
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close SaveChanges:=False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    Application.ScreenUpdating = True
    On Error GoTo 0

    ' 6. Итоговая сводка (не показываем, если проход прерван до начала)
    If Not aborted Then ПоказатьСводку updated, problems
    Exit Sub

Fail:
    Dim sErr As String
    sErr = Err.Description
    MsgBox "Ошибка при обновлении данных:" & vbCrLf & sErr, _
           vbCritical, "SmartCells"
    Resume CleanUp
End Sub

' ---------------------------------------------------------------------
'  Проход (а): поиск меток {адрес} или {Лист!адрес} и превращение их в CC.
'  Проблемные метки (кириллица в адресе, некорректный адрес) — подсветить
'  и включить в список, текст не трогать.
' ---------------------------------------------------------------------
Private Sub ПройтиПоТексту(ByVal doc As Document, ByVal problems As Collection)
    Dim rng As Range
    Dim inner As String, sheetName As String, addr As String, ref As String
    Dim cc As ContentControl

    Set rng = doc.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
        ' {  +  1 и более букв/цифр/пробелов/«!»/«_» (лат./кир.)  +  }
        ' допускается необязательное имя листа перед «!»: {Лист!A5}
        .Text = "\{[0-9A-Za-zА-Яа-яЁё !_]@\}"

        Do While .Execute
            ' rng указывает на найденную метку вида {A5} или {Лист!A5}
            inner = Mid$(rng.Text, 2, Len(rng.Text) - 2)   ' содержимое без скобок
            РазобратьСсылку inner, sheetName, addr

            If СодержитКириллицу(addr) Then
                ' Русские буквы в АДРЕСЕ — частая ошибка (имя листа может быть русским)
                ПодсветитьДиапазон rng, True
                Добавить problems, inner & " — адрес набран русскими буквами"
                rng.Collapse wdCollapseEnd

            ElseIf КорректныйАдрес(addr) Then
                ' Корректная метка — превращаем в Content Control.
                ' В тег пишем адрес в верхнем регистре, имя листа — как есть.
                If Len(sheetName) > 0 Then
                    ref = sheetName & "!" & UCase$(addr)
                Else
                    ref = UCase$(addr)
                End If
                Set cc = doc.ContentControls.Add(wdContentControlText, rng)
                cc.Tag = TAG_PREFIX & ref
                cc.Title = CC_TITLE
                ' Значение подставит проход (б); rng теперь — диапазон CC
                rng.SetRange cc.Range.End, cc.Range.End

            Else
                ' Некорректный формат адреса (например {5A})
                ПодсветитьДиапазон rng, True
                Добавить problems, inner & " — некорректный адрес"
                rng.Collapse wdCollapseEnd
            End If

            ' Продолжить поиск от текущей позиции до конца документа
            rng.End = doc.Content.End
        Loop
    End With
End Sub

' ---------------------------------------------------------------------
'  Проход (б): перечитать значения всех CC с тегом XL:* и обновить текст.
'  Лист выбирается по метке: {Лист!A5} — из указанного листа книги,
'  {A5} — из листа по умолчанию (из настроек). Возвращает число обновлённых.
' ---------------------------------------------------------------------
Private Function ПройтиПоCC(ByVal doc As Document, ByVal wb As Object, _
                            ByVal defaultWs As Object, ByVal problems As Collection) As Long
    Dim cc As ContentControl
    Dim ref As String, sheetName As String, addr As String
    Dim val As String, reason As String
    Dim ws As Object
    Dim code As Long
    Dim updated As Long

    For Each cc In doc.ContentControls
        If Left$(cc.Tag, Len(TAG_PREFIX)) = TAG_PREFIX Then
            ref = Mid$(cc.Tag, Len(TAG_PREFIX) + 1)
            РазобратьСсылку ref, sheetName, addr

            ' Выбрать лист: по умолчанию или указанный в метке
            If Len(sheetName) = 0 Then
                Set ws = defaultWs
            Else
                Set ws = НайтиЛист(wb, sheetName)
            End If

            If ws Is Nothing Then
                ' Указанный в метке лист не найден в книге
                ПодсветитьДиапазон cc.Range, True
                Добавить problems, ref & " — лист «" & sheetName & "» не найден"
            Else
                code = ПрочитатьЯчейку(ws, addr, val, reason)
                Select Case code
                    Case 0
                        ' Успех — подставить значение, снять подсветку
                        cc.Range.Text = val
                        ПодсветитьДиапазон cc.Range, False
                        updated = updated + 1
                    Case Else
                        ' Проблема — значение не затираем, подсвечиваем CC
                        ПодсветитьДиапазон cc.Range, True
                        Добавить problems, ref & " — " & reason
                End Select
            End If
        End If
    Next cc

    ПройтиПоCC = updated
End Function

' ---------------------------------------------------------------------
'  Разобрать ссылку «Лист!Адрес» на имя листа и адрес.
'  Без «!» — лист пустой (значит, лист по умолчанию).
' ---------------------------------------------------------------------
Private Sub РазобратьСсылку(ByVal s As String, ByRef sheetName As String, _
                            ByRef addr As String)
    Dim p As Long
    p = InStr(s, "!")
    If p > 0 Then
        sheetName = Trim$(Left$(s, p - 1))
        addr = Trim$(Mid$(s, p + 1))
    Else
        sheetName = ""
        addr = Trim$(s)
    End If
End Sub

' ---------------------------------------------------------------------
'  Чтение одной ячейки. Возврат: 0 — успех, иначе код проблемы.
'  out  — подставляемое значение (при code=0)
'  reason — причина проблемы (при code<>0)
' ---------------------------------------------------------------------
Private Function ПрочитатьЯчейку(ByVal ws As Object, ByVal addr As String, _
                                 ByRef out As String, ByRef reason As String) As Long
    Dim rng As Object
    Dim t As String

    On Error GoTo BadAddr
    Set rng = ws.Range(addr)        ' некорректный/вне сетки адрес -> ошибка
    On Error GoTo 0

    ' Ячейка с ошибкой Excel (#ДЕЛ/0! и т.п.)
    If IsError(rng.Value2) Then
        reason = "ошибка в ячейке (" & CStr(rng.Text) & ")"
        ПрочитатьЯчейку = 1
        Exit Function
    End If

    ' Значение — как видит пользователь; при узкой колонке (####) — из Value2
    t = CStr(rng.Text)
    If Len(t) > 0 And ТолькоРешётки(t) Then t = CStr(rng.Value2)

    ' Пустая ячейка — пустую строку не подставляем
    If IsEmpty(rng.Value2) Or Len(t) = 0 Then
        reason = "пустая ячейка"
        ПрочитатьЯчейку = 2
        Exit Function
    End If

    out = t
    ПрочитатьЯчейку = 0
    Exit Function

BadAddr:
    reason = "некорректный адрес или адрес вне листа"
    ПрочитатьЯчейку = 3
End Function

' =====================================================================
'  КОМАНДА: настройки (путь к файлу Excel и имя листа) -> реестр
' =====================================================================
Public Sub НастройкиSmartCells()
    Dim sPath As String, sSheet As String
    Dim curPath As String, curSheet As String

    ПрочитатьНастройки curPath, curSheet   ' текущие значения как подсказка

    ' Шаг 1. Путь к файлу
    sPath = Trim$(InputBox("Полный путь к файлу Excel (.xlsx):", _
                           "SmartCells — настройки (1/2)", curPath))
    If Len(sPath) = 0 Then Exit Sub        ' пользователь отменил
    If Len(Dir(sPath)) = 0 Then
        MsgBox "Файл не найден по указанному пути:" & vbCrLf & sPath, _
               vbExclamation, "SmartCells"
        Exit Sub
    End If

    ' Шаг 2. Имя листа
    sSheet = Trim$(InputBox("Имя листа в книге:", _
                            "SmartCells — настройки (2/2)", curSheet))
    If Len(sSheet) = 0 Then Exit Sub

    If Not ЛистСуществует(sPath, sSheet) Then
        MsgBox "Лист «" & sSheet & "» не найден в книге:" & vbCrLf & sPath, _
               vbExclamation, "SmartCells"
        Exit Sub
    End If

    ' Сохранить в реестр
    SaveSetting REG_APP, REG_SECTION, KEY_PATH, sPath
    SaveSetting REG_APP, REG_SECTION, KEY_SHEET, sSheet

    MsgBox "Настройки сохранены:" & vbCrLf & vbCrLf & _
           "Файл: " & sPath & vbCrLf & _
           "Лист: " & sSheet, vbInformation, "SmartCells"
End Sub

' =====================================================================
'  КОМАНДА: убрать связь — заменить CC на текущий текст (снять контролы)
' =====================================================================
Public Sub УбратьСвязьSmartCells()
    Dim ans As VbMsgBoxResult
    Dim scope As Range
    Dim cc As ContentControl
    Dim i As Long
    Dim removed As Long

    ans = MsgBox("Снять связь по всему документу?" & vbCrLf & vbCrLf & _
                 "Да — весь документ; Нет — только выделение.", _
                 vbQuestion + vbYesNoCancel, "SmartCells")
    If ans = vbCancel Then Exit Sub

    If ans = vbYes Then
        Set scope = ActiveDocument.Content
    Else
        Set scope = Selection.Range
    End If

    ' Идём с конца к началу, чтобы удаление не сбивало нумерацию
    For i = scope.ContentControls.Count To 1 Step -1
        Set cc = scope.ContentControls(i)
        If Left$(cc.Tag, Len(TAG_PREFIX)) = TAG_PREFIX Then
            cc.Delete False        ' убрать контрол, оставить текст
            removed = removed + 1
        End If
    Next i

    MsgBox "Связь снята. Обработано меток: " & removed, _
           vbInformation, "SmartCells"
End Sub

' =====================================================================
'  ВСПОМОГАТЕЛЬНЫЕ ПРОЦЕДУРЫ
' =====================================================================

' Прочитать настройки из реестра. True, если оба значения заданы.
Private Function ПрочитатьНастройки(ByRef sPath As String, _
                                    ByRef sSheet As String) As Boolean
    sPath = GetSetting(REG_APP, REG_SECTION, KEY_PATH, "")
    sSheet = GetSetting(REG_APP, REG_SECTION, KEY_SHEET, "")
    ПрочитатьНастройки = (Len(sPath) > 0 And Len(sSheet) > 0)
End Function

' Найти лист в уже открытой книге по имени (Nothing, если нет).
Private Function НайтиЛист(ByVal wb As Object, ByVal sSheet As String) As Object
    Dim ws As Object
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sSheet, vbTextCompare) = 0 Then
            Set НайтиЛист = ws
            Exit Function
        End If
    Next ws
    Set НайтиЛист = Nothing
End Function

' Проверить наличие листа в книге (открыть на чтение и закрыть).
Private Function ЛистСуществует(ByVal sPath As String, _
                                ByVal sSheet As String) As Boolean
    Dim xlApp As Object, xlWb As Object, ws As Object
    On Error GoTo Done
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    Set xlWb = xlApp.Workbooks.Open(FileName:=sPath, ReadOnly:=True, AddToMru:=False)
    Set ws = НайтиЛист(xlWb, sSheet)
    ЛистСуществует = Not (ws Is Nothing)
Done:
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close SaveChanges:=False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set ws = Nothing: Set xlWb = Nothing: Set xlApp = Nothing
End Function

' Проверка корректного адреса: 1–3 латинские буквы + 1–7 цифр.
Private Function КорректныйАдрес(ByVal s As String) As Boolean
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "^[A-Za-z]{1,3}[0-9]{1,7}$"
        re.IgnoreCase = True
        re.Global = False
    End If
    КорректныйАдрес = re.Test(s)
End Function

' Есть ли в строке хотя бы одна кириллическая буква.
Private Function СодержитКириллицу(ByVal s As String) As Boolean
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "[А-Яа-яЁё]"
        re.Global = False
    End If
    СодержитКириллицу = re.Test(s)
End Function

' Состоит ли строка целиком из символов "#" (узкая колонка Excel).
Private Function ТолькоРешётки(ByVal s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) <> "#" Then Exit Function
    Next i
    ТолькоРешётки = (Len(s) > 0)
End Function

' Подсветить/снять жёлтую заливку диапазона.
Private Sub ПодсветитьДиапазон(ByVal rng As Range, ByVal bOn As Boolean)
    On Error Resume Next
    If bOn Then
        rng.Shading.BackgroundPatternColor = CLR_YELLOW
    Else
        rng.Shading.BackgroundPatternColor = CLR_AUTO
    End If
End Sub

' Добавить строку в коллекцию проблем (без падения на дубликатах).
Private Sub Добавить(ByVal col As Collection, ByVal s As String)
    On Error Resume Next
    col.Add s
End Sub

' Показать итоговую сводку: обновлено / проблемы (первые 10).
Private Sub ПоказатьСводку(ByVal updated As Long, ByVal problems As Collection)
    Dim msg As String
    Dim i As Long, n As Long

    msg = "Обновлено ячеек: " & updated & vbCrLf
    n = problems.Count

    If n = 0 Then
        msg = msg & "Проблемных меток нет."
    Else
        msg = msg & "Проблемных меток: " & n & vbCrLf & vbCrLf
        For i = 1 To n
            If i > 10 Then
                msg = msg & "… и ещё " & (n - 10)
                Exit For
            End If
            msg = msg & "• " & problems(i) & vbCrLf
        Next i
    End If

    MsgBox msg, IIf(n = 0, vbInformation, vbExclamation), "SmartCells — итоги"
End Sub
