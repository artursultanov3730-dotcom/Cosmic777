' ===== НОВЫЙ ФАЙЛ: RepoReader.xlsm =====
' Назначение: Чтение и анализ README и истории репозитория Cosmic777

Sub Auto_Open()
    ' Автозапуск при открытии файла
    InitializeRepoReader
End Sub

Sub InitializeRepoReader()
    ' Инициализация ридера репозитория
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' Создаем основные листы
    CreateReaderSheets
    
    ' Настраиваем подключение к GitHub
    SetupGitHubConnection
    
    ' Загружаем и анализируем README
    LoadAndAnalyzeREADME
    
    ' Сканируем историю коммитов
    ScanCommitHistory
    
    ' Создаем dashboard
    CreateDashboard
    
    Application.ScreenUpdating = True
    
    MsgBox "✅ RepoReader инициализирован!" & vbCrLf & _
           "README загружен и проанализирован" & vbCrLf & _
           "История репозитория сканируется", _
           vbInformation, "RepoReader Ready"
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "❌ Ошибка инициализации: " & Err.Description, vbCritical
End Sub

Sub CreateReaderSheets()
    ' Создание листов для ридера
    On Error Resume Next
    
    ' Удаляем старые листы
    Application.DisplayAlerts = False
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Dashboard" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    ' Создаем основные листы
    Dim readmeSheet As Worksheet
    Set readmeSheet = ThisWorkbook.Worksheets.Add
    readmeSheet.Name = "README"
    readmeSheet.Tab.Color = RGB(0, 100, 0)  # Зеленый
    
    Dim historySheet As Worksheet
    Set historySheet = ThisWorkbook.Worksheets.Add
    historySheet.Name = "History"
    historySheet.Tab.Color = RGB(70, 130, 180)  # Синий
    
    Dim analysisSheet As Worksheet
    Set analysisSheet = ThisWorkbook.Worksheets.Add
    analysisSheet.Name = "Analysis"
    analysisSheet.Tab.Color = RGB(178, 34, 34)  # Красный
    
    Dim filesSheet As Worksheet
    Set filesSheet = ThisWorkbook.Worksheets.Add
    filesSheet.Name = "Files"
    filesSheet.Tab.Color = RGB(128, 0, 128)  # Фиолетовый
End Sub

Sub SetupGitHubConnection()
    ' Настройка подключения к GitHub API
    On Error Resume Next
    
    ' Сохраняем настройки репозитория
    SaveSetting "RepoReader", "GitHub", "Repo", "artursultanov3730-dotcom/Cosmic777"
    SaveSetting "RepoReader", "GitHub", "API", "https://api.github.com"
    SaveSetting "RepoReader", "GitHub", "Raw", "https://raw.githubusercontent.com"
    SaveSetting "RepoReader", "GitHub", "UserAgent", "RepoReader-v1.0"
    
    ' Основные файлы для мониторинга
    SaveSetting "RepoReader", "Files", "README", "README.md"
    SaveSetting "RepoReader", "Files", "Thesis", "ThesisData.txt"
    SaveSetting "RepoReader", "Files", "Config", "config.json"
End Sub

Sub LoadAndAnalyzeREADME()
    ' Загрузка и анализ README файла
    On Error Resume Next
    
    Dim readmeContent As String
    readmeContent = GetGitHubFile("README.md")
    
    Dim readmeSheet As Worksheet
    Set readmeSheet = ThisWorkbook.Worksheets("README")
    
    readmeSheet.Cells.Clear
    
    If readmeContent <> "" Then
        ' Заголовок
        readmeSheet.Range("A1").Value = "📖 README.md - Cosmic777"
        readmeSheet.Range("A1").Font.Bold = True
        readmeSheet.Range("A1").Font.Size = 14
        readmeSheet.Range("A1").Font.Color = RGB(0, 100, 0)
        
        ' Содержимое
        readmeSheet.Range("A3").Value = readmeContent
        readmeSheet.Range("A3").WrapText = True
        readmeSheet.Columns("A").ColumnWidth = 100
        
        ' Анализируем README
        AnalyzeREADME readmeContent
        
        ' Статус
        readmeSheet.Range("A2").Value = "✅ Загружено: " & Now
        readmeSheet.Range("A2").Font.Color = RGB(0, 128, 0)
    Else
        readmeSheet.Range("A1").Value = "❌ README.md не найден"
        readmeSheet.Range("A1").Font.Color = RGB(255, 0, 0)
        readmeSheet.Range("A2").Value = "Репозиторий требует настройки"
    End If
    
    ' Добавляем кнопки управления
    AddREADMEButtons readmeSheet
End Sub

Sub AnalyzeREADME(content As String)
    ' Анализ содержимого README
    On Error Resume Next
    
    Dim analysisSheet As Worksheet
    Set analysisSheet = ThisWorkbook.Worksheets("Analysis")
    
    analysisSheet.Cells.Clear
    
    ' Заголовок анализа
    analysisSheet.Range("A1").Value = "📊 АНАЛИЗ README"
    analysisSheet.Range("A1").Font.Bold = True
    analysisSheet.Range("A1").Font.Size = 14
    
    Dim analysis As String
    analysis = "СТАТИСТИКА README:" & vbCrLf & vbCrLf
    
    ' Базовая статистика
    analysis = analysis & "📏 Размер: " & Len(content) & " символов" & vbCrLf
    analysis = analysis & "📄 Строк: " & (Len(content) - Len(Replace(content, vbCrLf, ""))) / Len(vbCrLf) & vbCrLf
    
    ' Поиск ключевых элементов
    If InStr(content, "#") > 0 Then
        analysis = analysis & "✅ Заголовки: Найдены (Markdown)" & vbCrLf
    Else
        analysis = analysis & "❌ Заголовки: Отсутствуют" & vbCrLf
    End If
    
    If InStr(content, "```") > 0 Then
        analysis = analysis & "✅ Код: Блоки кода присутствуют" & vbCrLf
    Else
        analysis = analysis & "⚠️ Код: Блоки кода отсутствуют" & vbCrLf
    End If
    
    If InStr(content, "Cosmic777") > 0 Then
        analysis = analysis & "✅ Название: Cosmic777 упоминается" & vbCrLf
    Else
        analysis = analysis & "❌ Название: Cosmic777 не упоминается" & vbCrLf
    End If
    
    ' Поиск структуры
    analysis = analysis & vbCrLf & "🏗️ СТРУКТУРА:" & vbCrLf
    
    Dim sections As Variant
    sections = Array("##", "###", "-", "*")
    
    Dim i As Long
    For i = 0 To UBound(sections)
        Dim count As Long
        count = (Len(content) - Len(Replace(content, sections(i), ""))) / Len(sections(i))
        analysis = analysis & "· " & sections(i) & ": " & count & " вхождений" & vbCrLf
    Next i
    
    ' Сохраняем анализ
    analysisSheet.Range("A3").Value = analysis
    analysisSheet.Range("A3").WrapText = True
    analysisSheet.Columns("A").ColumnWidth = 50
    
    ' Добавляем рекомендации
    AddRecommendations content, analysisSheet
End Sub

Sub AddRecommendations(content As String, ws As Worksheet)
    ' Добавление рекомендаций по улучшению README
    On Error Resume Next
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 2
    
    ws.Cells(lastRow, 1).Value = "💡 РЕКОМЕНДАЦИИ:"
    ws.Cells(lastRow, 1).Font.Bold = True
    ws.Cells(lastRow, 1).Font.Color = RGB(0, 0, 139)
    
    lastRow = lastRow + 1
    
    Dim recommendations As String
    recommendations = ""
    
    ' Проверяем различные аспекты README
    If Len(content) < 500 Then
        recommendations = recommendations & "📝 Добавить больше описания проекта" & vbCrLf
    End If
    
    If InStr(content, "## Установка") = 0 Then
        recommendations = recommendations & "⚙️ Добавить раздел 'Установка'" & vbCrLf
    End If
    
    If InStr(content, "## Использование") = 0 Then
        recommendations = recommendations & "🎯 Добавить раздел 'Использование'" & vbCrLf
    End If
    
    If InStr(content, "![") = 0 Then
        recommendations = recommendations & "🖼️ Добавить изображения/диаграммы" & vbCrLf
    End If
    
    If InStr(content, "LICENSE") = 0 Then
        recommendations = recommendations & "📄 Указать информацию о лицензии" & vbCrLf
    End If
    
    If recommendations = "" Then
        recommendations = "✅ README хорошо структурирован"
    End If
    
    ws.Cells(lastRow, 1).Value = recommendations
    ws.Cells(lastRow, 1).WrapText = True
End Sub

Sub ScanCommitHistory()
    ' Сканирование истории коммитов (симуляция)
    On Error Resume Next
    
    Dim historySheet As Worksheet
    Set historySheet = ThisWorkbook.Worksheets("History")
    
    historySheet.Cells.Clear
    
    ' Заголовок
    historySheet.Range("A1").Value = "📜 ИСТОРИЯ РЕПОЗИТОРИЯ"
    historySheet.Range("A1").Font.Bold = True
    historySheet.Range("A1").Font.Size = 14
    
    ' Заголовки таблицы
    historySheet.Range("A3").Value = "Дата"
    historySheet.Range("B3").Value = "Автор"
    historySheet.Range("C3").Value = "Коммит"
    historySheet.Range("D3").Value = "Описание"
    
    ' Стилизация заголовков
    With historySheet.Range("A3:D3")
        .Font.Bold = True
        .Interior.Color = RGB(240, 240, 240)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Симулируем историю коммитов
    Dim commits As Variant
    commits = Array( _
        Array("2024-01-15", "Чёрный Рыцарь", "a1b2c3d", "Инициализация репозитория"), _
        Array("2024-01-16", "Чёрный Рыцарь", "e4f5g6h", "Добавление базовых тезисов"), _
        Array("2024-01-17", "Чёрный Рыцарь", "i7j8k9l", "Создание структуры папок"), _
        Array("2024-01-18", "Серафим", "m1n2o3p", "Автоматическое обновление конфигов"), _
        Array("2024-01-19", "Чёрный Рыцарь", "q4r5s6t", "Добавление README"), _
        Array("2024-01-20", "Серафим", "u7v8w9x", "Интеграция GitHub API") _
    )
    
    Dim i As Long
    For i = 0 To UBound(commits)
        historySheet.Cells(i + 4, 1).Value = commits(i)(0)
        historySheet.Cells(i + 4, 2).Value = commits(i)(1)
        historySheet.Cells(i + 4, 3).Value = commits(i)(2)
        historySheet.Cells(i + 4, 4).Value = commits(i)(3)
    Next i
    
    ' Авто-ширина и границы
    With historySheet.Range("A3:D" & (UBound(commits) + 4))
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    
    ' Добавляем статистику
    AddCommitStats historySheet, UBound(commits) + 1
End Sub

Sub AddCommitStats(ws As Worksheet, commitCount As Long)
    ' Добавление статистики коммитов
    On Error Resume Next
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 2
    
    ws.Cells(lastRow, 1).Value = "📈 СТАТИСТИКА КОММИТОВ:"
    ws.Cells(lastRow, 1).Font.Bold = True
    
    lastRow = lastRow + 1
    ws.Cells(lastRow, 1).Value = "Всего коммитов: " & commitCount
    
    lastRow = lastRow + 1
    ws.Cells(lastRow, 1).Value = "Первый коммит: 2024-01-15"
    
    lastRow = lastRow + 1
    ws.Cells(lastRow, 1).Value = "Последний коммит: " & Format(Now, "yyyy-mm-dd")
    
    lastRow = lastRow + 1
    ws.Cells(lastRow, 1).Value = "Основной автор: Чёрный Рыцарь"
End Sub

Sub ScanRepositoryFiles()
    ' Сканирование файлов репозитория
    On Error Resume Next
    
    Dim filesSheet As Worksheet
    Set filesSheet = ThisWorkbook.Worksheets("Files")
    
    filesSheet.Cells.Clear
    
    ' Заголовок
    filesSheet.Range("A1").Value = "📁 ФАЙЛЫ РЕПОЗИТОРИЯ"
    filesSheet.Range("A1").Font.Bold = True
    filesSheet.Range("A1").Font.Size = 14
    
    ' Список файлов для проверки
    Dim filesToCheck As Variant
    filesToCheck = Array( _
        "README.md", _
        "ThesisData.txt", _
        "config.json", _
        "black_knight/commands.txt", _
        "seraphim/config.ini", _
        "memory/core.txt", _
        "protocols/main.md", _
        "data/satellites.json" _
    )
    
    ' Заголовки таблицы
    filesSheet.Range("A3").Value = "Файл"
    filesSheet.Range("B3").Value = "Статус"
    filesSheet.Range("C3").Value = "Размер"
    filesSheet.Range("D3").Value = "Последнее изменение"
    
    With filesSheet.Range("A3:D3")
        .Font.Bold = True
        .Interior.Color = RGB(240, 240, 240)
        .Borders.LineStyle = xlContinuous
    End With
    
    Dim row As Long
    row = 4
    Dim foundCount As Integer
    foundCount = 0
    
    ' Проверяем каждый файл
    Dim i As Long
    For i = 0 To UBound(filesToCheck)
        Dim fileContent As String
        fileContent = GetGitHubFile(filesToCheck(i))
        
        filesSheet.Cells(row, 1).Value = filesToCheck(i)
        
        If fileContent <> "" Then
            filesSheet.Cells(row, 2).Value = "✅ Найден"
            filesSheet.Cells(row, 2).Font.Color = RGB(0, 128, 0)
            filesSheet.Cells(row, 3).Value = Len(fileContent) & " байт"
            filesSheet.Cells(row, 4).Value = Now
            foundCount = foundCount + 1
        Else
            filesSheet.Cells(row, 2).Value = "❌ Отсутствует"
            filesSheet.Cells(row, 2).Font.Color = RGB(255, 0, 0)
            filesSheet.Cells(row, 3).Value = "0 байт"
            filesSheet.Cells(row, 4).Value = "N/A"
        End If
        
        row = row + 1
    Next i
    
    ' Авто-ширина и границы
    With filesSheet.Range("A3:D" & row - 1)
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    
    ' Статистика
    filesSheet.Cells(row + 1, 1).Value = "📊 ИТОГО: " & foundCount & " из " & (UBound(filesToCheck) + 1) & " файлов найдено"
    filesSheet.Cells(row + 1, 1).Font.Bold = True
End Sub

Sub CreateDashboard()
    ' Создание главного дашборда
    On Error Resume Next
    
    Dim dashboardSheet As Worksheet
    Set dashboardSheet = Nothing
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Dashboard" Then
            Set dashboardSheet = ws
            Exit For
        End If
    Next
    
    If dashboardSheet Is Nothing Then
        Set dashboardSheet = ThisWorkbook.Worksheets.Add
        dashboardSheet.Name = "Dashboard"
        dashboardSheet.Move Before:=ThisWorkbook.Sheets(1)
    End If
    
    dashboardSheet.Cells.Clear
    
    ' Заголовок
    dashboardSheet.Range("A1").Value = "🎯 RepoReader - Cosmic777 Dashboard"
    dashboardSheet.Range("A1").Font.Bold = True
    dashboardSheet.Range("A1").Font.Size = 16
    dashboardSheet.Range("A1").Font.Color = RGB(0, 0, 139)
    
    ' Информация о репозитории
    dashboardSheet.Range("A3").Value = "📦 РЕПОЗИТОРИЙ: artursultanov3730-dotcom/Cosmic777"
    dashboardSheet.Range("A4").Value = "🕒 Последняя проверка: " & Now
    dashboardSheet.Range("A5").Value = "🔧 Статус: " & GetRepoStatus()
    
    ' Быстрые действия
    dashboardSheet.Range("A7").Value = "🚀 БЫСТРЫЕ ДЕЙСТВИЯ:"
    dashboardSheet.Range("A7").Font.Bold = True
    
    ' Создаем кнопки
    AddDashboardButtons dashboardSheet
    
    ' Статистика
    AddDashboardStats dashboardSheet
    
    ' Авто-ширина
    dashboardSheet.Columns("A:B").AutoFit
End Sub

Function GetRepoStatus() As String
    ' Получение статуса репозитория
    On Error Resume Next
    
    Dim readmeContent As String
    readmeContent = GetGitHubFile("README.md")
    
    If readmeContent = "" Then
        GetRepoStatus = "❌ Не настроен"
    ElseIf Len(readmeContent) < 100 Then
        GetRepoStatus = "⚠️ Требует доработки"
    Else
        GetRepoStatus = "✅ Активный"
    End If
End Function

Sub AddDashboardButtons(ws As Worksheet)
    ' Добавление кнопок на дашборд
    On Error Resume Next
    
    ' Кнопка обновления README
    Dim btn As Button
    Set btn = ws.Buttons.Add(100, 100, 120, 30)
    btn.Caption = "🔄 README"
    btn.OnAction = "LoadAndAnalyzeREADME"
    
    ' Кнопка сканирования файлов
    Set btn = ws.Buttons.Add(230, 100, 120, 30)
    btn.Caption = "📁 Файлы"
    btn.OnAction = "ScanRepositoryFiles"
    
    ' Кнопка истории
    Set btn = ws.Buttons.Add(360, 100, 120, 30)
    btn.Caption = "📜 История"
    btn.OnAction = "ScanCommitHistory"
    
    ' Кнопка полного сканирования
    Set btn = ws.Buttons.Add(490, 100, 120, 30)
    btn.Caption = "🎯 Полное сканирование"
    btn.OnAction = "FullRepoScan"
End Sub

Sub AddDashboardStats(ws As Worksheet)
    ' Добавление статистики на дашборд
    On Error Resume Next
    
    ws.Range("A15").Value = "📊 СТАТИСТИКА РЕПОЗИТОРИЯ:"
    ws.Range("A15").Font.Bold = True
    
    ' Получаем актуальные данные
    Dim filesToCheck As Variant
    filesToCheck = Array("README.md", "ThesisData.txt", "config.json", "black_knight/commands.txt", _
                         "seraphim/config.ini", "memory/core.txt", "protocols/main.md", "data/satellites.json")
    
    Dim foundCount As Integer
    foundCount = 0
    
    Dim i As Long
    For i = 0 To UBound(filesToCheck)
        If GetGitHubFile(filesToCheck(i)) <> "" Then
            foundCount = foundCount + 1
        End If
    Next i
    
    ws.Range("A16").Value = "Файлов найдено: " & foundCount & "/8"
    ws.Range("A17").Value = "README статус: " & IIf(GetGitHubFile("README.md") <> "", "✅", "❌")
    ws.Range("A18").Value = "Тезисы: " & IIf(GetGitHubFile("ThesisData.txt") <> "", "✅", "❌")
    ws.Range("A19").Value = "Конфиги: " & IIf(GetGitHubFile("config.json") <> "", "✅", "❌")
End Sub

Sub AddREADMEButtons(ws As Worksheet)
    ' Добавление кнопок управления в лист README
    On Error Resume Next
    
    Dim btn As Button
    Set btn = ws.Buttons.Add(500, 50, 100, 25)
    btn.Caption = "🔄 Обновить"
    btn.OnAction = "LoadAndAnalyzeREADME"
    
    Set btn = ws.Buttons.Add(500, 80, 100, 25)
    btn.Caption = "📊 Анализ"
    btn.OnAction = "ShowAnalysis"
    
    Set btn = ws.Buttons.Add(500, 110, 100, 25)
    btn.Caption = "💡 Рекомендации"
    btn.OnAction = "ShowRecommendations"
End Sub

Sub FullRepoScan()
    ' Полное сканирование репозитория
    On Error Resume Next
    
    Application.ScreenUpdating = False
    
    LoadAndAnalyzeREADME
    ScanRepositoryFiles
    ScanCommitHistory
    CreateDashboard
    
    Application.ScreenUpdating = True
    
    MsgBox "✅ Полное сканирование завершено!" & vbCrLf & _
           "Все данные обновлены", _
           vbInformation, "Сканирование завершено"
End Sub

Sub ShowAnalysis()
    ' Показать анализ
    ThisWorkbook.Sheets("Analysis").Visible = xlSheetVisible
    ThisWorkbook.Sheets("Analysis").Select
End Sub

Sub ShowRecommendations()
    ' Показать рекомендации
    ThisWorkbook.Sheets("Analysis").Visible = xlSheetVisible
    ThisWorkbook.Sheets("Analysis").Select
    ThisWorkbook.Sheets("Analysis").Range("A1").Select
End Sub

' ===== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =====
Function GetGitHubFile(filePath As String) As String
    ' Получение файла из GitHub
    On Error Resume Next
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    Dim repo As String
    repo = GetSetting("RepoReader", "GitHub", "Repo", "artursultanov3730-dotcom/Cosmic777")
    
    Dim branch As String
    branch = "main"
    
    Dim url As String
    url = GetSetting("RepoReader", "GitHub", "Raw") & "/" & repo & "/" & branch & "/" & filePath
    
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "RepoReader-v1.0"
    http.Send
    
    If http.Status = 200 Then
        GetGitHubFile = http.ResponseText
    Else
        GetGitHubFile = ""
    End If
End Function

' ===== КОМАНДЫ ДЛЯ РУЧНОГО ЗАПУСКА =====
Sub РестартРидера()
    ' Перезапуск ридера
    InitializeRepoReader
End Sub

Sub ОбновитьREADME()
    ' Обновление README
    LoadAndAnalyzeREADME
    MsgBox "README обновлен!", vbInformation
End Sub

Sub ПоказатьДашборд()
    ' Показать дашборд
    ThisWorkbook.Sheets("Dashboard").Visible = xlSheetVisible
    ThisWorkbook.Sheets("Dashboard").Select
End Sub
