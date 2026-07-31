Attribute VB_Name = "SmartCells"
' =====================================================================
'  SmartCells — подстановка значений ячеек Excel в текст документа Word
'  Основной модуль: команды, разбор меток, работа с Excel, сводка.
'  Механика: метки {адрес} или {Лист!адрес} превращаются в Content
'  Control с тегом "XL:ссылка"; повторное обновление идёт по этим CC.
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

' --- Прочие константы ---
Private Const XL_CALC_AUTO As Long = -4105         ' xlCalculationAutomatic
' «Пароль-заглушка»: для файла без пароля игнорируется, для защищённого
' даёт перехватываемую ошибку вместо невидимого модального диалога Excel
Private Const PWD_STUB As String = "#SmartCellsNoPwd#"

' =====================================================================
'  ГЛАВНАЯ КОМАНДА: обновить все значения в активном документе
' =====================================================================
Public Sub ОбновитьДанные()
    ' Защита: команда доступна с панели и без открытых документов
    If Documents.Count = 0 Then
        MsgBox "Нет открытого документа.", vbExclamation, "SmartCells"
        Exit Sub
    End If
    ОбновитьДанныеДок ActiveDocument
End Sub

' ---------------------------------------------------------------------
'  Обновление конкретного документа (вызывается также из AppEvents
'  при событии DocumentOpen — там документ передаётся явно).
' ---------------------------------------------------------------------
Public Sub ОбновитьДанныеДок(ByVal doc As Document)
    Dim sPath As String, sSheet As String
    Dim xlApp As Object, xlWb As Object, xlWs As Object
    Dim updated As Long
    Dim sFatal As String                 ' текст фатальной ошибки (показ после закрытия Excel)
    Dim problems As Collection
    Set problems = New Collection

    ' 1. Прочитать настройки, при отсутствии — сначала диалог настроек
    If Not ПрочитатьНастройки(sPath, sSheet) Then
        НастройкиSmartCells
        If Not ПрочитатьНастройки(sPath, sSheet) Then
            MsgBox "Настройки не заданы. Обновление отменено.", _
                   vbExclamation, "SmartCells"
            Exit Sub
        End If
    End If

    ' 2. Проверить наличие файла Excel — если нет, документ не трогаем вообще
    If Not ФайлСуществует(sPath) Then
        MsgBox "Файл Excel не найден по пути:" & vbCrLf & vbCrLf & _
               sPath & vbCrLf & vbCrLf & _
               "Проверьте путь через команду «НастройкиSmartCells».", _
               vbExclamation, "SmartCells"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' 3. Открыть Excel один раз на весь проход (late binding, ReadOnly, невидимо)
    On Error GoTo Fail
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    Set xlWb = xlApp.Workbooks.Open(FileName:=sPath, ReadOnly:=True, _
                                    AddToMru:=False, UpdateLinks:=0, _
                                    Password:=PWD_STUB)

    ' Включить автопересчёт: если книга сохранена в ручном режиме,
    ' формулы могли остаться с несвежими значениями
    On Error Resume Next
    xlApp.Calculation = XL_CALC_AUTO
    On Error GoTo Fail

    ' Найти лист по умолчанию; если листа нет — документ не меняем
    Set xlWs = НайтиЛист(xlWb, sSheet)
    If xlWs Is Nothing Then
        sFatal = "Лист «" & sSheet & "» не найден в книге:" & vbCrLf & _
                 sPath & vbCrLf & vbCrLf & _
                 "Проверьте имя листа через «НастройкиSmartCells»."
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

    ' 6. Сообщения — только ПОСЛЕ закрытия Excel (чтобы модальное окно
    '    не держало невидимый процесс EXCEL.EXE)
    If Len(sFatal) > 0 Then
        MsgBox sFatal, vbExclamation, "SmartCells"
    Else
        ПоказатьСводку updated, problems
    End If
    Exit Sub

Fail:
    sFatal = "Ошибка при обновлении данных:" & vbCrLf & Err.Description
    Resume CleanUp
End Sub

' ---------------------------------------------------------------------
'  Проход (а): поиск меток {адрес} или {Лист!адрес} и превращение их в CC.
'  Совпадения внутри уже существующих CC пропускаются (их обновит проход
'  (б) по тегу) — иначе Word падает на попытке вложенного контрола.
'  Обычный текст в скобках ({Итого}, {примечание 1}) меткой не считается
'  и не трогается. Проблемные метки — подсветить и включить в список.
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
            ' Совпадение внутри существующего контрола (например, метка,
            ' оставшаяся в CC после проблемной ячейки) — пропустить:
            ' вложенные CC запрещены, а этот CC обработает проход (б)
            If Not rng.ParentContentControl Is Nothing Then
                rng.Collapse wdCollapseEnd
                rng.End = doc.Content.End
            Else
                ' rng указывает на найденную метку вида {A5} или {Лист!A5}
                inner = Mid$(rng.Text, 2, Len(rng.Text) - 2)   ' содержимое без скобок
                РазобратьСсылку inner, sheetName, addr

                If Not ПохожеНаМетку(sheetName, addr) Then
                    ' Обычный текст в фигурных скобках — не метка, не трогаем
                    rng.Collapse wdCollapseEnd

                ElseIf СодержитКириллицу(addr) Then
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
            End If
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
    Dim sVal As String, reason As String
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
                code = ПрочитатьЯчейку(ws, addr, sVal, reason)
                Select Case code
                    Case 0
                        ' Успех — подставить значение, снять подсветку
                        cc.Range.Text = sVal
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
'  Похоже ли содержимое скобок на метку (а не на обычный текст).
'  Метка-кандидат: указан лист ({Лист!A5}), либо короткая строка без
'  пробелов и подчёркиваний, где есть и буквы, и цифры ({A5}, {А5}, {5A}).
'  Всё остальное ({Итого}, {примечание 1}) — обычный текст, пропускаем.
' ---------------------------------------------------------------------
Private Function ПохожеНаМетку(ByVal sheetName As String, ByVal addr As String) As Boolean
    Dim i As Long, ch As String
    Dim hasLetter As Boolean, hasDigit As Boolean

    If Len(sheetName) > 0 Then      ' есть «!» — пользователь явно писал метку
        ПохожеНаМетку = True
        Exit Function
    End If
    If Len(addr) = 0 Or Len(addr) > 11 Then Exit Function
    If InStr(addr, " ") > 0 Or InStr(addr, "_") > 0 Then Exit Function

    For i = 1 To Len(addr)
        ch = Mid$(addr, i, 1)
        If ch >= "0" And ch <= "9" Then hasDigit = True Else hasLetter = True
    Next i
    ПохожеНаМетку = hasLetter And hasDigit
End Function

' ---------------------------------------------------------------------
'  Чтение одной ячейки. Возврат: 0 — успех, иначе код проблемы.
'  sOut — подставляемое значение (при code=0)
'  reason — причина проблемы (при code<>0)
' ---------------------------------------------------------------------
Private Function ПрочитатьЯчейку(ByVal ws As Object, ByVal addr As String, _
                                 ByRef sOut As String, ByRef reason As String) As Long
    Dim rng As Object
    Dim t As String

    On Error GoTo BadAddr
    Set rng = ws.Range(addr)        ' некорректный/вне сетки адрес -> ошибка
    On Error GoTo ReadFail          ' дальше любые сбои чтения — код 5

    ' Диапазон вместо одиночной ячейки (возможно при ручной правке тега)
    If rng.Cells.Count > 1 Then
        reason = "метка ссылается на диапазон, а не на одну ячейку"
        ПрочитатьЯчейку = 4
        Exit Function
    End If

    ' Ячейка с ошибкой Excel (#ДЕЛ/0! и т.п.)
    If IsError(rng.Value2) Then
        reason = "ошибка в ячейке (" & CStr(rng.Text) & ")"
        ПрочитатьЯчейку = 1
        Exit Function
    End If

    ' Значение — как видит пользователь (rng.Text)
    t = CStr(rng.Text)

    ' Узкая колонка («####»): расширить колонку в памяти и перечитать —
    ' fallback на Value2 сразу нельзя: для даты он вернул бы «45870»
    ' вместо «15.07.2026». Книга открыта ReadOnly — на диск не пишется.
    If Len(t) > 0 And ТолькоРешётки(t) Then
        On Error Resume Next
        rng.EntireColumn.AutoFit
        On Error GoTo ReadFail
        t = CStr(rng.Text)
        If ТолькоРешётки(t) Then t = CStr(rng.Value2)   ' последний резерв
    End If

    ' Пустая ячейка — пустую строку не подставляем
    If IsEmpty(rng.Value2) Or Len(t) = 0 Then
        reason = "пустая ячейка"
        ПрочитатьЯчейку = 2
        Exit Function
    End If

    ' Переводы строк (Alt+Enter в Excel) заменяем пробелами:
    ' plain-text CC без MultiLine не принимает текст с переводом строки
    t = Replace(t, vbCrLf, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbCr, " ")

    sOut = t
    ПрочитатьЯчейку = 0
    Exit Function

BadAddr:
    reason = "некорректный адрес или адрес вне листа"
    ПрочитатьЯчейку = 3
    Exit Function

ReadFail:
    reason = "не удалось прочитать ячейку (" & Err.Description & ")"
    ПрочитатьЯчейку = 5
End Function

' =====================================================================
'  КОМАНДА: настройки (путь к файлу Excel и имя листа) -> реестр
' =====================================================================
Public Sub НастройкиSmartCells()
    Dim sPath As String, sSheet As String
    Dim curPath As String, curSheet As String
    Dim reason As String

    ПрочитатьНастройки curPath, curSheet   ' текущие значения как подсказка

    ' Шаг 1. Путь к файлу
    sPath = ОчиститьПуть(InputBox("Полный путь к файлу Excel (.xlsx):", _
                                  "SmartCells — настройки (1/2)", curPath))
    If Len(sPath) = 0 Then Exit Sub        ' пользователь отменил
    If Not ФайлСуществует(sPath) Then
        MsgBox "Файл не найден по указанному пути:" & vbCrLf & sPath, _
               vbExclamation, "SmartCells"
        Exit Sub
    End If

    ' Шаг 2. Имя листа
    sSheet = Trim$(InputBox("Имя листа в книге:", _
                            "SmartCells — настройки (2/2)", curSheet))
    If Len(sSheet) = 0 Then Exit Sub

    ' Проверка: книга открывается и лист существует (с внятной причиной)
    If Not ПроверитьКнигуИЛист(sPath, sSheet, reason) Then
        MsgBox "Проверка не пройдена:" & vbCrLf & reason & vbCrLf & vbCrLf & _
               "Файл: " & sPath, vbExclamation, "SmartCells"
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

    ' Защита: команда доступна с панели и без открытых документов
    If Documents.Count = 0 Then
        MsgBox "Нет открытого документа.", vbExclamation, "SmartCells"
        Exit Sub
    End If

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

    ' Курсор стоит ВНУТРИ метки-CC (выделение не охватывает контрол
    ' целиком) — снять и этот контрол тоже
    If ans = vbNo Then
        Set cc = Selection.Range.ParentContentControl
        If Not cc Is Nothing Then
            If Left$(cc.Tag, Len(TAG_PREFIX)) = TAG_PREFIX Then
                cc.Delete False
                removed = removed + 1
            End If
        End If
    End If

    MsgBox "Связь снята. Обработано меток: " & removed, _
           vbInformation, "SmartCells"
End Sub

' =====================================================================
'  ВСПОМОГАТЕЛЬНЫЕ ПРОЦЕДУРЫ
' =====================================================================

' Прочистить путь из буфера обмена: убрать кавычки («Копировать как путь»
' в Проводнике добавляет их) и лишние пробелы.
Private Function ОчиститьПуть(ByVal s As String) As String
    ОчиститьПуть = Trim$(Replace(s, Chr$(34), ""))
End Function

' Существует ли файл (защищено от ошибок Dir на кривых путях/отключённых дисках).
Private Function ФайлСуществует(ByVal sPath As String) As Boolean
    On Error Resume Next
    ФайлСуществует = (Len(Dir(sPath)) > 0)
    On Error GoTo 0
End Function

' Прочитать настройки из реестра. True, если оба значения заданы.
Private Function ПрочитатьНастройки(ByRef sPath As String, _
                                    ByRef sSheet As String) As Boolean
    sPath = ОчиститьПуть(GetSetting(REG_APP, REG_SECTION, KEY_PATH, ""))
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

' Проверить, что книга открывается и лист существует.
' Различает причины: «книга не открылась» и «листа нет».
Private Function ПроверитьКнигуИЛист(ByVal sPath As String, ByVal sSheet As String, _
                                     ByRef reason As String) As Boolean
    Dim xlApp As Object, xlWb As Object, ws As Object
    On Error Resume Next

    Set xlApp = CreateObject("Excel.Application")
    If xlApp Is Nothing Then
        reason = "не удалось запустить Excel"
        GoTo Done
    End If
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Err.Clear
    Set xlWb = xlApp.Workbooks.Open(FileName:=sPath, ReadOnly:=True, _
                                    AddToMru:=False, UpdateLinks:=0, _
                                    Password:=PWD_STUB)
    If xlWb Is Nothing Then
        reason = "не удалось открыть книгу (файл занят, повреждён или защищён паролем)"
        GoTo Done
    End If

    Set ws = НайтиЛист(xlWb, sSheet)
    If ws Is Nothing Then
        reason = "в книге нет листа «" & sSheet & "»"
    Else
        ПроверитьКнигуИЛист = True
    End If

Done:
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
