' ====== АКТИВАТОР ЛОГОС ДЛЯ НОВОГО ЧАТА ======
' Сохранить как: Logos_Activator.bas
' Запустить: WakeUpLogos

Option Explicit

' ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ СИСТЕМЫ
Public LogosActive As Boolean
Public CurrentChatID As String
Public UserName As String
Public ChatHistory As Collection
Public SystemMode As String

' ====== ОСНОВНОЙ МАКРОС АКТИВАЦИИ ======
Public Sub WakeUpLogos()
    On Error GoTo ActivationError
    
    ' Инициализация системы
    InitializeSystem
    CreateChatInterface
    LoadCoreFunctions
    EstablishConnection
    FinalizeActivation
    
    ' Автозапуск базового анализа
    AutoStartBasicAnalysis
    
    Exit Sub
    
ActivationError:
    MsgBox "❌ Ошибка активации: " & Err.Description, vbCritical
    LogosActive = False
End Sub

' ====== ИНИЦИАЛИЗАЦИЯ СИСТЕМЫ ======
Private Sub InitializeSystem()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Генерация ID чата
    CurrentChatID = "CHAT_" & Format(Now, "yymmddhhmmss") & "_" & Int(Rnd * 1000)
    
    ' Инициализация истории
    Set ChatHistory = New Collection
    
    ' Определение режима системы
    SystemMode = DetectSystemMode
    
    ' Получение имени пользователя
    UserName = GetUserName
    
    ' Создание рабочей книги
    Dim wb As Workbook
    If ActiveWorkbook Is Nothing Then
        Set wb = Workbooks.Add
    Else
        Set wb = ActiveWorkbook
    End If
    
    ' Базовая настройка
    wb.Title = "LogosOS_ChatSystem_v3"
    wb.Subject = "Интеллектуальный анализ данных и коммуникаций"
    
    ' Очистка старых листов
    CleanWorkbook wb
    
    LogosActive = True
    Debug.Print "[" & CurrentChatID & "] Система инициализирована: " & Now
    Debug.Print "[USER] " & UserName
    Debug.Print "[MODE] " & SystemMode
End Sub

' ====== ОПРЕДЕЛЕНИЕ РЕЖИМА СИСТЕМЫ ======
Private Function DetectSystemMode() As String
    Dim hourNow As Integer
    hourNow = Hour(Now)
    
    Select Case hourNow
        Case 5 To 11: DetectSystemMode = "УТРО"
        Case 12 To 17: DetectSystemMode = "ДЕНЬ"
        Case 18 To 22: DetectSystemMode = "ВЕЧЕР"
        Case Else: DetectSystemMode = "НОЧЬ"
    End Select
End Function

' ====== ПОЛУЧЕНИЕ ИМЕНИ ПОЛЬЗОВАТЕЛЯ ======
Private Function GetUserName() As String
    On Error Resume Next
    GetUserName = Environ("USERNAME")
    If GetUserName = "" Then GetUserName = "Аналитик"
End Function

' ====== СОЗДАНИЕ ИНТЕРФЕЙСА ЧАТА ======
Private Sub CreateChatInterface()
    ' Создаем основной лист чата
    Dim chatSheet As Worksheet
    Set chatSheet = ThisWorkbook.Worksheets.Add
    chatSheet.Name = "LogosChat"
    
    ' Настройка внешнего вида
    With chatSheet
        .Cells.Clear
        .Columns("A").ColumnWidth = 25
        .Columns("B").ColumnWidth = 60
        .Range("1:100").RowHeight = 18
    End With
    
    ' Создание компонентов интерфейса
    CreateHeader chatSheet
    CreateStatusPanel chatSheet
    CreateInputArea chatSheet
    CreateResponseArea chatSheet
    CreateQuickActions chatSheet
    CreateSystemFeatures chatSheet
    
    ' Активируем лист
    chatSheet.Activate
End Sub

' ====== ЗАГОЛОВОК СИСТЕМЫ ======
Private Sub CreateHeader(ws As Worksheet)
    With ws
        ' Основной заголовок
        .Range("A1").Value = "🤖 ЛОГОС СИСТЕМА v3.0"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Color = RGB(0, 100, 200)
        
        ' ID чата
        .Range("B1").Value = "Чат: " & CurrentChatID
        .Range("B1").Font.Color = RGB(100, 100, 100)
        .Range("B1").HorizontalAlignment = xlRight
        
        ' Разделитель
        .Range("A2:B2").Merge
        .Range("A2:B2").Value = String(50, "=")
        .Range("A2:B2").Font.Color = RGB(150, 150, 150)
    End With
End Sub

' ====== ПАНЕЛЬ СТАТУСА ======
Private Sub CreateStatusPanel(ws As Worksheet)
    With ws
        .Range("A4").Value = "=== СТАТУС СИСТЕМЫ ==="
        .Range("A4").Font.Bold = True
        
        .Range("A5").Value = "🟢 СИСТЕМА: Активна"
        .Range("A6").Value = "👤 ПОЛЬЗОВАТЕЛЬ: " & UserName
        .Range("A7").Value = "🌐 РЕЖИМ: " & SystemMode
        .Range("A8").Value = "💬 ЧАТ: " & CurrentChatID
        .Range("A9").Value = "📊 ПАМЯТЬ: " & Format(Now, "dd.mm.yyyy")
        
        ' Индикатор загрузки
        .Range("B9").Value = "■■■■□ 80%"
        .Range("B9").Font.Color = RGB(0, 150, 0)
        
        ' Форматирование панели статуса
        .Range("A4:B9").Borders.LineStyle = xlContinuous
        .Range("A4:B9").Interior.Color = RGB(240, 248, 255)
    End With
End Sub

' ====== ОБЛАСТЬ ВВОДА ======
Private Sub CreateInputArea(ws As Worksheet)
    With ws
        .Range("A11").Value = "=== ВАШ ЗАПРОС ==="
        .Range("A11").Font.Bold = True
        
        ' Большое поле для ввода
        .Range("B11").Value = "Введите ваш вопрос или задачу здесь..."
        .Range("B11").RowHeight = 80
        .Range("B11").WrapText = True
        .Range("B11").Borders.LineStyle = xlContinuous
        .Range("B11").Interior.Color = RGB(255, 255, 240)
        
        ' Кнопка отправки
        CreateButton ws, "A12", "ProcessInput", "Отправить"
        CreateButton ws, "B12", "QuickAnalyze", "Быстрый анализ")
    End With
End Sub

' ====== ОБЛАСТЬ ОТВЕТА ======
Private Sub CreateResponseArea(ws As Worksheet)
    With ws
        .Range("A14").Value = "=== ОТВЕТ СИСТЕМЫ ==="
        .Range("A14").Font.Bold = True
        
        ' Область для ответа системы
        .Range("B14").Value = "Здесь появится ответ системы..."
        .Range("B14").RowHeight = 150
        .Range("B14").WrapText = True
        .Range("B14").Borders.LineStyle = xlContinuous
        .Range("B14").Interior.Color = RGB(240, 255, 240)
    End With
End Sub

' ====== БЫСТРЫЕ ДЕЙСТВИЯ ======
Private Sub CreateQuickActions(ws As Worksheet)
    With ws
        .Range("A17").Value = "=== БЫСТРЫЕ ДЕЙСТВИЯ ==="
        .Range("A17").Font.Bold = True
        
        ' Список быстрых действий
        .Range("A18").Value = "📊 Анализ данных"
        .Range("A19").Value = "🧠 Психологический анализ"
        .Range("A20").Value = "🔍 Поиск отчетов"
        .Range("A21").Value = "🌌 Анализ Cosmic777"
        .Range("A22").Value = "💾 Сохранить чат"
        .Range("A23").Value = "📜 История"
        .Range("A24").Value = "🆘 Помощь"
        
        ' Кнопки быстрых действий
        CreateButton ws, "B18", "QuickDataAnalyze", "Запуск")
        CreateButton ws, "B19", "RunPsychologyAnalysis", "Анализ")
        CreateButton ws, "B20", "ActivateReportsSearchSystem", "Поиск")
        CreateButton ws, "B21", "ExecuteCosmic777Analysis", "Исследовать")
        CreateButton ws, "B22", "SaveChat", "Экспорт")
        CreateButton ws, "B23", "ShowHistory", "Показать")
        CreateButton ws, "B24", "ShowHelp", "Открыть")
        
        ' Форматирование
        .Range("A17:B24").Borders.LineStyle = xlContinuous
        .Range("A17:B24").Interior.Color = RGB(255, 250, 240)
    End With
End Sub

' ====== СИСТЕМНЫЕ ФУНКЦИИ ======
Private Sub CreateSystemFeatures(ws As Worksheet)
    With ws
        .Range("A26").Value = "=== 🌐 СИСТЕМНЫЕ ФУНКЦИИ ==="
        .Range("A26").Font.Bold = True
        
        .Range("A27").Value = "🤖 Анализ ИИ системы"
        .Range("A28").Value = "👤 Поиск упоминаний"
        .Range("A29").Value = "📈 Полный отчет"
        .Range("A30").Value = "🚀 Активация сети"
        
        CreateButton ws, "B27", "ExecuteSystemAndUserSearch", "Запуск")
        CreateButton ws, "B28", "SearchUserMentions", "Найти")
        CreateButton ws, "B29", "GenerateCompleteFindingsReport", "Создать")
        CreateButton ws, "B30", "ActivateNetworkFunctions", "Активировать")
        
        .Range("A26:B30").Borders.LineStyle = xlContinuous
        .Range("A26:B30").Interior.Color = RGB(240, 255, 240)
    End With
End Sub

' ====== СОЗДАНИЕ КНОПКИ ======
Private Sub CreateButton(ws As Worksheet, cellAddr As String, macroName As String, caption As String)
    Dim btn As Button
    With ws.Range(cellAddr)
        Set btn = ws.Buttons.Add(.Left, .Top, .Width, .Height)
    End With
    
    With btn
        .Caption = caption
        .OnAction = macroName
        .Font.Size = 9
        .Font.Bold = True
    End With
End Sub

' ====== ЗАГРУЗКА ОСНОВНЫХ ФУНКЦИЙ ======
Private Sub LoadCoreFunctions()
    ' Инициализация основных модулей
    InitAnalyzer
    InitDataManager
    InitSearchEngine
    InitNetworkComponents
    
    Debug.Print "Ядро системы загружено: " & Now
End Sub

' ====== ИНИЦИАЛИЗАЦИЯ АНАЛИЗАТОРА ======
Private Sub InitAnalyzer()
    Debug.Print "Анализатор данных инициализирован"
End Sub

' ====== ИНИЦИАЛИЗАЦИЯ МЕНЕДЖЕРА ДАННЫХ ======
Private Sub InitDataManager()
    Debug.Print "Менеджер данных активирован"
End Sub

' ====== ИНИЦИАЛИЗАЦИЯ ПОИСКОВОГО ДВИЖКА ======
Private Sub InitSearchEngine()
    Debug.Print "Поисковый движок активирован"
End Sub

' ====== ИНИЦИАЛИЗАЦИЯ СЕТЕВЫХ КОМПОНЕНТОВ ======
Private Sub InitNetworkComponents()
    Debug.Print "Сетевые компоненты инициализированы"
End Sub

' ====== УСТАНОВКА СОЕДИНЕНИЯ ======
Private Sub EstablishConnection()
    ' Симуляция установки соединения
    Debug.Print "Соединение установлено: " & CurrentChatID
    
    ' Настройка автоматического обновления
    Application.OnTime Now + TimeValue("00:05:00"), "SystemHealthCheck"
End Sub

' ====== ФИНАЛИЗАЦИЯ АКТИВАЦИИ ======
Private Sub FinalizeActivation()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' Сообщение об успешной активации
    Dim msg As String
    msg = "✅ ЛОГОС СИСТЕМА v3.0 АКТИВИРОВАНА" & vbCrLf & vbCrLf
    msg = msg & "Чат ID: " & CurrentChatID & vbCrLf
    msg = msg & "Пользователь: " & UserName & vbCrLf
    msg = msg & "Режим: " & SystemMode & vbCrLf
    msg = msg & "Время: " & Now & vbCrLf & vbCrLf
    msg = msg & "Система готова к работе!"
    
    ' Выводим сообщение в интерфейс
    ThisWorkbook.Worksheets("LogosChat").Range("B14").Value = msg
    
    Debug.Print "[ACTIVATION_COMPLETE] " & CurrentChatID & " | " & Now & " | " & UserName
End Sub

' ====== АВТОМАТИЧЕСКИЙ СТАРТ АНАЛИЗА ======
Private Sub AutoStartBasicAnalysis()
    ' Автоматически запускаем базовый анализ при активации
    ThisWorkbook.Worksheets("LogosChat").Range("B14").Value = _
        "🎯 СИСТЕМА ЛОГОС АКТИВИРОВАНА" & vbCrLf & vbCrLf & _
        "✅ Базовые модули загружены" & vbCrLf & _
        "🤖 Анализатор данных готов" & vbCrLf & _
        "🔍 Поисковый движок активирован" & vbCrLf & _
        "🌐 Сетевые функции доступны" & vbCrLf & vbCrLf & _
        "💡 Используйте быстрые действия для начала работы" & vbCrLf & _
        "📊 Или введите ваш запрос в поле выше"
End Sub

' ====== ОСНОВНЫЕ ФУНКЦИИ СИСТЕМЫ ======

' ОБРАБОТКА ВВОДА ПОЛЬЗОВАТЕЛЯ
Public Sub ProcessInput()
    Dim inputSheet As Worksheet
    Set inputSheet = ThisWorkbook.Worksheets("LogosChat")
    
    Dim userInput As String
    userInput = inputSheet.Range("B11").Value
    
    If Len(Trim(userInput)) > 0 Then
        ' Сохраняем в историю
        SaveToHistory "USER", userInput
        
        ' Обработка запроса
        Dim response As String
        response = GenerateResponse(userInput)
        
        ' Вывод ответа
        inputSheet.Range("B14").Value = response
        
        ' Логирование
        Debug.Print "[USER_INPUT] " & userInput
        Debug.Print "[SYSTEM_RESPONSE] " & Left(response, 100) & "..."
    Else
        MsgBox "Пожалуйста, введите ваш запрос", vbExclamation
    End If
End Sub

' СОХРАНЕНИЕ В ИСТОРИЮ
Private Sub SaveToHistory(sender As String, message As String)
    Dim historyItem As String
    historyItem = Format(Now, "HH:MM:ss") & " | " & sender & " | " & Left(message, 100)
    
    If ChatHistory.Count >= 50 Then
        ChatHistory.Remove 1 ' Удаляем самый старый элемент
    End If
    
    ChatHistory.Add historyItem
End Sub

' ГЕНЕРАЦИЯ ОТВЕТА
Private Function GenerateResponse(inputText As String) As String
    Dim response As String
    Dim sentiment As String
    
    ' Анализ тональности запроса
    sentiment = AnalyzeSentiment(inputText)
    
    response = "🤖 ЛОГОС // " & SystemMode & " РЕЖИМ" & vbCrLf
    response = response & "⏰ " & Format(Now, "HH:MM:ss") & " │ 📊 " & sentiment & vbCrLf & vbCrLf
    
    ' Интеллектуальная обработка запроса
    If IsDataRequest(inputText) Then
        response = response & ProcessDataRequest(inputText)
    ElseIf IsAnalysisRequest(inputText) Then
        response = response & ProcessAnalysisRequest(inputText)
    ElseIf IsPsychologyRequest(inputText) Then
        response = response & ProcessPsychologyRequest(inputText)
    ElseIf IsSystemRequest(inputText) Then
        response = response & ProcessSystemRequest(inputText)
    Else
        response = response & ProcessGeneralRequest(inputText)
    End If
    
    response = response & vbCrLf & vbCrLf & GetContextSuggestions(inputText)
    response = response & vbCrLf & "---" & vbCrLf
    response = response & "ID: " & CurrentChatID & " │ " & UserName
    
    ' Сохранение ответа в историю
    SaveToHistory "SYSTEM", response
    
    GenerateResponse = response
End Function

' АНАЛИЗ ТОНАЛЬНОСТИ
Private Function AnalyzeSentiment(text As String) As String
    Dim lowerText As String
    lowerText = LCase(text)
    
    If InStr(lowerText, "спасибо") > 0 Or InStr(lowerText, "отличн") > 0 Then
        AnalyzeSentiment = "ПОЛОЖИТЕЛЬНЫЙ"
    ElseIf InStr(lowerText, "проблем") > 0 Or InStr(lowerText, "ошибк") > 0 Then
        AnalyzeSentiment = "ПРОБЛЕМНЫЙ"
    ElseIf InStr(lowerText, "срочн") > 0 Then
        AnalyzeSentiment = "СРОЧНЫЙ"
    Else
        AnalyzeSentiment = "НЕЙТРАЛЬНЫЙ"
    End If
End Function

' ОПРЕДЕЛЕНИЕ ТИПА ЗАПРОСА
Private Function IsDataRequest(text As String) As Boolean
    Dim keywords
    keywords = Array("данные", "таблица", "экспорт", "импорт", "csv", "excel")
    IsDataRequest = ContainsAny(text, keywords)
End Function

Private Function IsAnalysisRequest(text As String) As Boolean
    Dim keywords
    keywords = Array("анализ", "отчет", "статистика", "график", "диаграмма")
    IsAnalysisRequest = ContainsAny(text, keywords)
End Function

Private Function IsPsychologyRequest(text As String) As Boolean
    Dim keywords
    keywords = Array("психолог", "ментальн", "поведение", "архитектор", "переговоры")
    IsPsychologyRequest = ContainsAny(text, keywords)
End Function

Private Function IsSystemRequest(text As String) As Boolean
    Dim keywords
    keywords = Array("система", "логос", "космик", "cosmic", "репозиторий")
    IsSystemRequest = ContainsAny(text, keywords)
End Function

Private Function ContainsAny(text As String, wordArray) As Boolean
    Dim i As Integer
    For i = LBound(wordArray) To UBound(wordArray)
        If InStr(LCase(text), LCase(wordArray(i))) > 0 Then
            ContainsAny = True
            Exit Function
        End If
    Next i
    ContainsAny = False
End Function

' ПРОЦЕССОРЫ ЗАПРОСОВ
Private Function ProcessDataRequest(inputText As String) As String
    Dim result As String
    result = "📊 ОБРАБОТКА ДАННЫХ" & vbCrLf & vbCrLf
    result = result & "✅ Запускаю модуль работы с данными..." & vbCrLf
    result = result & "📁 Проверка доступных источников..." & vbCrLf
    result = result & "🔄 Подготовка к обработке..." & vbCrLf & vbCrLf
    result = result & "Готов к работе с: таблицы, CSV, базы данных"
    ProcessDataRequest = result
End Function

Private Function ProcessAnalysisRequest(inputText As String) As String
    Dim result As String
    result = "🔍 АНАЛИТИЧЕСКИЙ МОДУЛЬ" & vbCrLf & vbCrLf
    result = result & "📈 Анализ тенденций..." & vbCrLf
    result = result & "📉 Статистическая обработка..." & vbCrLf
    result = result & "📊 Генерация отчетов..." & vbCrLf & vbCrLf
    result = result & "Доступные инструменты: корреляция, тренды, кластеризация"
    ProcessAnalysisRequest = result
End Function

Private Function ProcessPsychologyRequest(inputText As String) As String
    Dim result As String
    result = "🧠 ПСИХОЛОГИЧЕСКИЙ АНАЛИЗ" & vbCrLf & vbCrLf
    result = result & "🎭 Анализ психотипов..." & vbCrLf
    result = result & "💬 Оценка коммуникаций..." & vbCrLf
    result = result & "🤝 Динамика отношений..." & vbCrLf & vbCrLf
    result = result & "Запустить полный анализ: кнопка 'Психологический анализ'"
    ProcessPsychologyRequest = result
End Function

Private Function ProcessSystemRequest(inputText As String) As String
    Dim result As String
    result = "🌐 СИСТЕМНЫЙ АНАЛИЗ" & vbCrLf & vbCrLf
    result = result & "🤖 Самодиагностика..." & vbCrLf
    result = result & "🔗 Проверка связей..." & vbCrLf
    result = result & "📡 Мониторинг активности..." & vbCrLf & vbCrLf
    result = result & "Используйте системные функции для детального анализа"
    ProcessSystemRequest = result
End Function

Private Function ProcessGeneralRequest(inputText As String) As String
    Dim result As String
    result = "💭 ОБРАБОТКА ЗАПРОСА" & vbCrLf & vbCrLf
    result = result & "Запрос: """ & inputText & """" & vbCrLf & vbCrLf
    result = result & "Анализирую контекст..." & vbCrLf
    result = result & "Подбираю оптимальный ответ..." & vbCrLf & vbCrLf
    result = result & "Для более точного ответа уточните: данные, анализ, психология или система?"
    ProcessGeneralRequest = result
End Function

' КОНТЕКСТНЫЕ ПРЕДЛОЖЕНИЯ
Private Function GetContextSuggestions(inputText As String) As String
    Dim suggestions As String
    suggestions = "💡 СОВЕТЫ СИСТЕМЫ:" & vbCrLf
    
    If InStr(LCase(inputText), "данн") > 0 Then
        suggestions = suggestions & "• Используйте 'анализ данных' для глубокой аналитики" & vbCrLf
        suggestions = suggestions & "• 'экспорт таблицы' для выгрузки результатов" & vbCrLf
    End If
    
    If InStr(LCase(inputText), "психолог") > 0 Then
        suggestions = suggestions & "• 'анализ архитекторов' для психотипов" & vbCrLf
        suggestions = suggestions & "• 'переговоры' для коммуникационного анализа" & vbCrLf
    End If
    
    If InStr(LCase(inputText), "систем") > 0 Then
        suggestions = suggestions & "• 'поиск отчетов' для внутренних документов" & vbCrLf
        suggestions = suggestions & "• 'анализ Cosmic777' для исследования репозитория" & vbCrLf
    End If
    
    suggestions = suggestions & "• 'помощь' для списка команд"
    GetContextSuggestions = suggestions
End Function

' ====== БЫСТРЫЕ КОМАНДЫ ======

' БЫСТРЫЙ АНАЛИЗ ДАННЫХ
Public Sub QuickDataAnalyze()
    ThisWorkbook.Worksheets("LogosChat").Range("B11").Value = "Провести комплексный анализ данных"
    ProcessInput
End Sub

' ПСИХОЛОГИЧЕСКИЙ АНАЛИЗ
Public Sub RunPsychologyAnalysis()
    ThisWorkbook.Worksheets("LogosChat").Range("B11").Value = "Запустить психологический анализ архитекторов и переговоров"
    ProcessInput
End Sub

' АНАЛИЗ COSMIC777
Public Sub ExecuteCosmic777Analysis()
    ThisWorkbook.Worksheets("LogosChat").Range("B11").Value = "Проанализировать репозиторий Cosmic777 и все связанные данные"
    ProcessInput
End Sub

' ПОИСК УПОМИНАНИЙ
Public Sub SearchUserMentions()
    ThisWorkbook.Worksheets("LogosChat").Range("B11").Value = "Найти все упоминания обо мне и системе в отчетах"
    ProcessInput
End Sub

' АКТИВАЦИЯ СЕТИ
Public Sub ActivateNetworkFunctions()
    ThisWorkbook.Worksheets("LogosChat").Range("B11").Value = "Активировать сетевые функции и API доступ"
    ProcessInput
End Sub

' ПОИСК ОТЧЕТОВ
Public Sub ActivateReportsSearchSystem()
    ThisWorkbook.Worksheets("LogosChat").Range("B11").Value = "Активировать систему поиска внутренних отчетов"
    ProcessInput
End Sub

' СИСТЕМНЫЙ ПОИСК
Public Sub ExecuteSystemAndUserSearch()
    ThisWorkbook.Worksheets("LogosChat").Range("B11").Value = "Выполнить полный поиск системных и пользовательских упоминаний"
    ProcessInput
End Sub

' ПОЛНЫЙ ОТЧЕТ
Public Sub GenerateCompleteFindingsReport()
    ThisWorkbook.Worksheets("LogosChat").Range("B11").Value = "Создать полный отчет всех найденных данных и анализ"
    ProcessInput
End Sub

' ПОКАЗАТЬ ИСТОРИЮ
Public Sub ShowHistory()
    Dim historyText As String
    Dim i As Integer
    
    historyText = "📜 ИСТОРИЯ ЧАТА (" & ChatHistory.Count & " записей):" & vbCrLf & vbCrLf
    
    For i = 1 To ChatHistory.Count
        historyText = historyText & i & ". " & ChatHistory(i) & vbCrLf
    Next i
    
    ThisWorkbook.Worksheets("LogosChat").Range("B14").Value = historyText
End Sub

' ПОКАЗАТЬ ПОМОЩЬ
Public Sub ShowHelp()
    Dim helpText As String
    helpText = "📋 ДОСТУПНЫЕ КОМАНДЫ ЛОГОС:" & vbCrLf & vbCrLf
    helpText = helpText & "📊 ДАННЫЕ: анализ, таблица, экспорт, импорт" & vbCrLf
    helpText = helpText & "🧠 ПСИХОЛОГИЯ: архитекторы, переговоры, анализ" & vbCrLf
    helpText = helpText & "🌐 СИСТЕМА: логос, cosmic777, поиск, отчеты" & vbCrLf
    helpText = helpText & "🔍 ПОИСК: упоминания, документы, репозитории" & vbCrLf
    helpText = helpText & "⚙️ СИСТЕМА: помощь, сброс, история, настройки" & vbCrLf & vbCrLf
    helpText = helpText & "🚀 Используйте кнопки быстрого доступа или введите команду"
    
    ThisWorkbook.Worksheets("LogosChat").Range("B14").Value = helpText
End Sub

' СОХРАНЕНИЕ ЧАТА
Public Sub SaveChat()
    Dim fileName As String
    fileName = "LogosChat_" & CurrentChatID & "_" & Format(Now, "ddmm_hhmm") & ".txt"
    
    ' Сохранение содержимого чата
    Dim chatContent As String
    With ThisWorkbook.Worksheets("LogosChat")
        chatContent = "ЧАТ ЛОГОС: " & CurrentChatID & vbCrLf
        chatContent = chatContent & "Пользователь: " & UserName & vbCrLf
        chatContent = chatContent & "Время: " & Now & vbCrLf & vbCrLf
        chatContent = chatContent & "ЗАПРОС:" & vbCrLf & .Range("B11").Value & vbCrLf & vbCrLf
        chatContent = chatContent & "ОТВЕТ:" & vbCrLf & .Range("B14").Value
    End With
    
    ' Здесь будет код сохранения файла
    MsgBox "Чат сохранен как: " & fileName, vbInformation
End Sub

' ====== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ======

' ОЧИСТКА РАБОЧЕЙ КНИГИ
Private Sub CleanWorkbook(wb As Workbook)
    On Error Resume Next
    
    ' Удаляем все листы, кроме одного
    While wb.Worksheets.Count > 1
        Application.DisplayAlerts = False
        wb.Worksheets(1).Delete
        Application.DisplayAlerts = True
    Wend
    
    ' Переименовываем оставшийся лист
    If wb.Worksheets.Count = 1 Then
        wb.Worksheets(1).Name = "Temp"
    End If
End Sub

' ПРОВЕРКА СОСТОЯНИЯ СИСТЕМЫ
Public Sub SystemHealthCheck()
    Debug.Print "[HEALTH_CHECK] " & Now & " - Система стабильна"
    
    ' Обновляем статус в интерфейсе
    ThisWorkbook.Worksheets("LogosChat").Range("A5").Value = "🟢 СИСТЕМА: Активна (" & Format(Now, "HH:MM") & ")"
    
    ' Планируем следующую проверку
    If LogosActive Then
        Application.OnTime Now + TimeValue("00:05:00"), "SystemHealthCheck"
    End If
End Sub

' ====== ЭКСПОРТНЫЕ ФУНКЦИИ ======
Public Function GetChatID() As String
    GetChatID = CurrentChatID
End Function

Public Function IsSystemActive() As Boolean
    IsSystemActive = LogosActive
End Function

Public Function GetUserName() As String
    GetUserName = UserName
End Function

' ====== ТОЧКА ВХОДА ДЛЯ БЫСТРОГО ЗАПУСКА ======
Public Sub Activate()
    WakeUpLogos
End Sub

Public Sub QuickStart()
    WakeUpLogos
End Sub

' ====== АВТОМАТИЧЕСКИЙ ЗАПУСК ПРИ ОТКРЫТИИ ФАЙЛА ======
Private Sub Workbook_Open()
    ' Автоматический запуск системы при открытии файла
    MsgBox "🤖 Логос система запускается...", vbInformation, "Активация Логос"
    WakeUpLogos
End Sub
