package main

import (
    "fmt"
    "io"
    "net/http"
    "sync"
    "time"

    "fyne.io/fyne/v2"
    "fyne.io/fyne/v2/app"
    "fyne.io/fyne/v2/container"
    "fyne.io/fyne/v2/dialog"
    "fyne.io/fyne/v2/storage"
    "fyne.io/fyne/v2/widget"

    "github.com/xuri/excelize/v2"
)

// Структура для хранения информации об URL
type urlInfo struct {
    Row int
    URL string
}

// Ошибка проверки
type urlError struct {
    Row int
    URL string
    Err string
}

var userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"

// checkURL выполняет HEAD-запрос (при необходимости GET) и возвращает ошибку или пустую строку
func checkURL(row int, url string, treatRedirect bool) urlError {
    client := &http.Client{Timeout: 8 * time.Second}

    doRequest := func(method string) (*http.Response, error) {
        req, _ := http.NewRequest(method, url, nil)
        req.Header.Set("User-Agent", userAgent)
        return client.Do(req)
    }

    // HEAD запрос без следования редиректам
    resp, err := doRequest("HEAD")
    if err != nil {
        return urlError{row, url, fmt.Sprintf("Connection Error: %v", err)}
    }
    defer resp.Body.Close()

    // Если метод не разрешён (403, 405, 501) — пробуем GET
    if resp.StatusCode == http.StatusForbidden || resp.StatusCode == http.StatusMethodNotAllowed || resp.StatusCode == http.StatusNotImplemented {
        resp2, err2 := doRequest("GET")
        if err2 != nil {
            return urlError{row, url, fmt.Sprintf("Connection Error: %v", err2)}
        }
        defer resp2.Body.Close()
        resp = resp2
    }

    // Анализируем статус
    if resp.StatusCode < 400 {
        if treatRedirect && resp.StatusCode >= 300 && resp.StatusCode < 400 {
            loc := resp.Header.Get("Location")
            return urlError{row, url, fmt.Sprintf("Redirect %d → %s", resp.StatusCode, loc)}
        }
        // Если редирект и мы разрешаем следовать, проверяем конечный статус
        if resp.StatusCode >= 300 && resp.StatusCode < 400 && !treatRedirect {
            // Делаем запрос с автоматическим переходом по редиректам
            finalResp, finalErr := client.Get(url)
            if finalErr != nil {
                return urlError{row, url, fmt.Sprintf("Connection Error after redirect: %v", finalErr)}
            }
            defer finalResp.Body.Close()
            if finalResp.StatusCode < 400 {
                return urlError{Row: row, URL: url} // нет ошибки
            }
            return urlError{row, url, fmt.Sprintf("HTTP %d %s", finalResp.StatusCode, finalResp.Status)}
        }
        return urlError{Row: row, URL: url} // всё хорошо
    }
    return urlError{row, url, fmt.Sprintf("HTTP %d %s", resp.StatusCode, resp.Status)}
}

// runChecks параллельно проверяет список URL и возвращает только ошибочные
func runChecks(urls []urlInfo, treatRedirect bool, progress chan<- float64) []urlError {
    var wg sync.WaitGroup
    results := make(chan urlError, len(urls))
    total := len(urls)

    for i, u := range urls {
        wg.Add(1)
        go func(idx int, info urlInfo) {
            defer wg.Done()
            res := checkURL(info.Row, info.URL, treatRedirect)
            if res.Err != "" {
                results <- res
            }
            progress <- float64(idx+1) / float64(total) * 100
        }(i, u)
    }

    go func() {
        wg.Wait()
        close(results)
        close(progress)
    }()

    var errors []urlError
    for e := range results {
        errors = append(errors, e)
    }
    return errors
}

// loadSheetsAndColumns открывает xlsx и возвращает список листов и столбцов первого листа
func loadSheetsAndColumns(path string) ([]string, []string, error) {
    f, err := excelize.OpenFile(path)
    if err != nil {
        return nil, nil, err
    }
    defer f.Close()

    sheets := f.GetSheetList()
    if len(sheets) == 0 {
        return nil, nil, fmt.Errorf("нет листов в файле")
    }
    // читаем первую строку как заголовки
    rows, err := f.GetRows(sheets[0])
    if err != nil || len(rows) == 0 {
        return sheets, []string{"A"}, nil // fallback
    }
    headers := rows[0]
    return sheets, headers, nil
}

// extractURLs извлекает URL из указанного листа и столбца
func extractURLs(path, sheet, colLetter string) ([]urlInfo, error) {
    f, err := excelize.OpenFile(path)
    if err != nil {
        return nil, err
    }
    defer f.Close()

    rows, err := f.GetRows(sheet)
    if err != nil {
        return nil, err
    }
    if len(rows) < 2 {
        return nil, fmt.Errorf("нет данных")
    }

    // Преобразуем букву столбца в индекс (A->0, B->1, ...)
    colIdx := int(colLetter[0] - 'A')
    if colIdx < 0 || colIdx >= len(rows[0]) {
        return nil, fmt.Errorf("столбец %s не найден", colLetter)
    }

    var urls []urlInfo
    for i := 1; i < len(rows); i++ { // со второй строки (индекс 1)
        row := rows[i]
        if len(row) <= colIdx {
            continue
        }
        cellValue := row[colIdx]
        // Пытаемся получить гиперссылку (excelize хранит их отдельно)
        cellRef := fmt.Sprintf("%s%d", colLetter, i+1) // строки в excelize с 1
        if link, ok := f.GetCellHyperLink(sheet, cellRef); ok && link.Link != "" {
            urls = append(urls, urlInfo{Row: i + 1, URL: link.Link})
        } else if cellValue != "" {
            urls = append(urls, urlInfo{Row: i + 1, URL: cellValue})
        }
    }
    return urls, nil
}

// saveReport сохраняет ошибки в xlsx рядом с исходным файлом
func saveReport(sourcePath string, errors []urlError) (string, error) {
    f := excelize.NewFile()
    sheet := "Ошибки"
    f.SetSheetName("Sheet1", sheet)

    // Заголовки
    f.SetCellValue(sheet, "A1", "Строка в Excel")
    f.SetCellValue(sheet, "B1", "URL")
    f.SetCellValue(sheet, "C1", "Тип ошибки")

    // Заливка красным
    styleRed, _ := f.NewStyle(&excelize.Style{
        Fill: excelize.Fill{Type: "pattern", Color: []string{"FFCCCC"}, Pattern: 1},
    })

    for i, e := range errors {
        row := i + 2
        f.SetCellValue(sheet, fmt.Sprintf("A%d", row), e.Row)
        f.SetCellValue(sheet, fmt.Sprintf("B%d", row), e.URL)
        f.SetCellValue(sheet, fmt.Sprintf("C%d", row), e.Err)
        f.SetCellStyle(sheet, fmt.Sprintf("A%d", row), fmt.Sprintf("C%d", row), styleRed)
    }

    // Сохраняем в ту же папку, что и исходный файл
    dir := filepath.Dir(sourcePath)
    reportPath := filepath.Join(dir, "report_bad_urls.xlsx")
    if err := f.SaveAs(reportPath); err != nil {
        return "", err
    }
    return reportPath, nil
}

func main() {
    myApp := app.New()
    myWindow := myApp.NewWindow("Проверка URL из Excel")
    myWindow.Resize(fyne.NewSize(520, 420))

    // Переменные состояния
    var (
        filePath      string
        sheetNames    []string
        columnHeaders []string
        selectedSheet string
        selectedCol   string
    )

    // Элементы GUI
    fileEntry := widget.NewEntry()
    fileEntry.Disable()
    fileBtn := widget.NewButton("Обзор", func() {
        fd := dialog.NewFileOpen(func(reader fyne.URIReadCloser, err error) {
            if err != nil || reader == nil {
                return
            }
            filePath = reader.URI().Path()
            fileEntry.SetText(filePath)
            reader.Close()

            // Загружаем листы и столбцы
            s, c, err := loadSheetsAndColumns(filePath)
            if err != nil {
                dialog.ShowError(err, myWindow)
                return
            }
            sheetNames = s
            columnHeaders = c
            sheetSelector.Options = s
            if len(s) > 0 {
                sheetSelector.SetSelected(s[0])
                selectedSheet = s[0]
            }
            colSelector.Options = c
            if len(c) > 0 {
                colSelector.SetSelected(c[0])
                selectedCol = c[0]
            }
        }, myWindow)
        fd.SetFilter(storage.NewExtensionFileFilter([]string{".xlsx"}))
        fd.Show()
    })

    sheetSelector := widget.NewSelect([]string{}, func(s string) {
        selectedSheet = s
    })
    colSelector := widget.NewSelect([]string{}, func(s string) {
        selectedCol = s
    })

    treatRedirectCheck := widget.NewCheck("Считать редиректы ошибкой", nil)

    progressBar := widget.NewProgressBar()
    statusLabel := widget.NewLabel("")

    checkBtn := widget.NewButton("Проверить URL", func() {
        if filePath == "" {
            dialog.ShowInformation("Ошибка", "Сначала выберите Excel-файл", myWindow)
            return
        }
        checkBtn.Disable()
        progressBar.SetValue(0)
        statusLabel.SetText("Проверка...")

        // Определяем букву столбца по имени (если имя не буква, ищем индекс)
        colLetter := selectedCol
        if len(colLetter) > 1 || colLetter < "A" || colLetter > "Z" {
            // Ищем индекс столбца в заголовках
            for i, h := range columnHeaders {
                if h == selectedCol {
                    colLetter = string(rune('A' + i))
                    break
                }
            }
        }

        go func() {
            urls, err := extractURLs(filePath, selectedSheet, colLetter)
            if err != nil {
                checkBtn.Enable()
                statusLabel.SetText("Ошибка чтения файла")
                dialog.ShowError(err, myWindow)
                return
            }
            if len(urls) == 0 {
                checkBtn.Enable()
                statusLabel.SetText("Нет URL для проверки")
                dialog.ShowInformation("Результат", "В столбце нет URL", myWindow)
                return
            }

            progressChan := make(chan float64, len(urls))
            treatRedirect := treatRedirectCheck.Checked

            var errors []urlError
            go func() {
                errors = runChecks(urls, treatRedirect, progressChan)
            }()

            for p := range progressChan {
                progressBar.SetValue(p)
            }
            checkBtn.Enable()
            if len(errors) == 0 {
                statusLabel.SetText("Готово")
                dialog.ShowInformation("Отлично!", "Все URL работают корректно", myWindow)
                return
            }

            reportPath, err := saveReport(filePath, errors)
            if err != nil {
                statusLabel.SetText("Ошибка сохранения отчёта")
                dialog.ShowError(err, myWindow)
                return
            }
            statusLabel.SetText("Готово")
            dialog.ShowInformation("Отчёт сохранён", fmt.Sprintf("Найдено %d проблемных URL.\nФайл: %s", len(errors), reportPath), myWindow)
        }()
    })

    infoBtn := widget.NewButton("Инструкция", func() {
        text := `ИНСТРУКЦИЯ
1. Нажмите «Обзор» и выберите Excel-файл (.xlsx).
2. Выберите лист и столбец, где находятся URL.
   Поддерживаются явные ссылки и гиперссылки.
3. При необходимости отметьте редиректы как ошибки.
4. Нажмите «Проверить URL».
5. Отчёт будет сохранён в папку с исходным файлом.`
        dialog.ShowInformation("Инструкция", text, myWindow)
    })

    // Компоновка интерфейса
    content := container.NewVBox(
        widget.NewLabel("1. Выберите Excel-файл"),
        container.NewBorder(nil, nil, nil, fileBtn, fileEntry),
        widget.NewLabel("2. Выберите лист"),
        sheetSelector,
        widget.NewLabel("3. Выберите столбец с URL"),
        colSelector,
        widget.NewLabel("Настройки проверки"),
        treatRedirectCheck,
        container.NewHBox(checkBtn, infoBtn),
        progressBar,
        statusLabel,
    )

    myWindow.SetContent(content)
    myWindow.ShowAndRun()
}
