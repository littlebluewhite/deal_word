package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"sync"
	"sync/atomic"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/widget"
	"github.com/ledongthuc/pdf"
	"github.com/pdfcpu/pdfcpu/pkg/api"
)

func main() {
	// 建立 Fyne 應用程式
	a := app.New()
	w := a.NewWindow("PDF 分割器")

	// 進度條和狀態標籤
	progressBar := widget.NewProgressBar()
	statusLabel := widget.NewLabel("請選擇 PDF 檔案")
	errorList := widget.NewMultiLineEntry()
	errorList.SetPlaceHolder("錯誤訊息將顯示在此處...")
	errorList.Disable() // 設置為不可編輯

	// 選擇輸出資料夾的按鈕和輸入框
	outputDirEntry := widget.NewEntry()
	outputDirEntry.SetPlaceHolder("選擇輸出資料夾")
	outputDirEntry.Disable()
	selectOutputDirButton := widget.NewButton("選擇輸出資料夾", func() {
		// 打開資料夾選擇對話框
		dialog.ShowFolderOpen(func(uri fyne.ListableURI, err error) {
			if err != nil {
				dialog.ShowError(err, w)
				return
			}
			if uri == nil {
				// 使用者取消了資料夾選擇
				return
			}
			outputDirEntry.SetText(uri.Path())
		}, w)
	})

	// 建立取消處理的按鈕
	cancelProcessing := make(chan struct{})
	cancelButton := widget.NewButton("取消", func() {
		// 發送取消信號
		close(cancelProcessing)
	})
	cancelButton.Disable() // 初始時禁用

	// 提前宣告 selectFileButton
	var selectFileButton *widget.Button

	// 建立選擇檔案的按鈕
	selectFileButton = widget.NewButton("選擇 PDF 檔案並開始處理", func() {
		// 打開檔案選擇對話框
		dialog.ShowFileOpen(func(reader fyne.URIReadCloser, err error) {
			if err != nil {
				dialog.ShowError(err, w)
				return
			}
			if reader == nil {
				// 使用者取消了檔案選擇
				return
			}
			inputPath := reader.URI().Path()
			reader.Close() // 關閉檔案

			// 檢查輸出資料夾是否已選擇
			outputDir := outputDirEntry.Text
			if outputDir == "" {
				dialog.ShowError(fmt.Errorf("請先選擇輸出資料夾"), w)
				return
			}

			// 啟用取消按鈕，禁用選擇檔案按鈕
			cancelButton.Enable()
			selectFileButton.Disable()

			// 開始處理 PDF
			cancelProcessing = make(chan struct{}) // 重置取消通道
			go processPDF(inputPath, outputDir, w, progressBar, statusLabel, errorList, cancelProcessing, selectFileButton, cancelButton)
		}, w)
	})

	// 佈局調整

	// 使 outputDirEntry 根據視窗寬度拉伸
	outputDirContainer := container.NewBorder(nil, nil, nil, selectOutputDirButton, outputDirEntry)

	// 上部控件
	topContainer := container.NewVBox(
		statusLabel,
		progressBar,
		outputDirContainer,
		selectFileButton,
		cancelButton,
		widget.NewLabel("錯誤訊息："),
	)

	// 將 errorList 放入可滾動的容器中
	errorListScroll := container.NewVScroll(errorList)

	// 主容器
	content := container.NewBorder(
		topContainer,    // Top
		nil,             // Bottom
		nil,             // Left
		nil,             // Right
		errorListScroll, // Center
	)

	// 設定視窗內容並顯示
	w.SetContent(content)
	w.Resize(fyne.NewSize(600, 400))
	w.ShowAndRun()
}

// 處理 PDF 檔案
func processPDF(inputPath string, outputDir string, w fyne.Window, progressBar *widget.ProgressBar, statusLabel *widget.Label, errorList *widget.Entry, cancelProcessing chan struct{}, selectFileButton, cancelButton *widget.Button) {
	// 確保在處理完畢後，恢復按鈕狀態
	defer func() {
		selectFileButton.Enable()
		cancelButton.Disable()
	}()

	// 檢查檔案是否存在
	if _, err := os.Stat(inputPath); os.IsNotExist(err) {
		dialog.ShowError(fmt.Errorf("檔案 %s 不存在", inputPath), w)
		return
	}

	// 建立輸出資料夾
	err := createOutputDir(outputDir)
	if err != nil {
		dialog.ShowError(fmt.Errorf("無法建立輸出資料夾: %v", err), w)
		return
	}

	// 更新狀態
	statusLabel.SetText("正在拆分 PDF 檔案...")
	progressBar.SetValue(0)
	errorList.SetText("")

	// 拆分 PDF 檔案
	err = splitPDF(inputPath, outputDir)
	if err != nil {
		dialog.ShowError(fmt.Errorf("無法拆分 PDF 檔案: %v", err), w)
		return
	}

	// 處理拆分後的 PDF 頁面
	err = processPDFPages(outputDir, inputPath, w, progressBar, statusLabel, errorList, cancelProcessing)
	if err != nil {
		dialog.ShowError(fmt.Errorf("處理 PDF 頁面時出錯: %v", err), w)
		return
	}

	progressBar.SetValue(1)
	statusLabel.SetText("處理完成")
	dialog.ShowInformation("完成", fmt.Sprintf("已將 %s 拆分並重命名到資料夾 %s 中。", inputPath, outputDir), w)
}

// 建立輸出資料夾
func createOutputDir(outputDir string) error {
	if _, err := os.Stat(outputDir); os.IsNotExist(err) {
		err := os.MkdirAll(outputDir, 0755)
		if err != nil {
			return err
		}
	}
	return nil
}

// 拆分 PDF 檔案
func splitPDF(inputPath, outputDir string) error {
	err := api.SplitFile(inputPath, outputDir, 1, nil)
	return err
}

// 處理拆分後的 PDF 頁面
func processPDFPages(outputDir, inputPath string, w fyne.Window, progressBar *widget.ProgressBar, statusLabel *widget.Label, errorList *widget.Entry, cancelProcessing chan struct{}) error {
	files, err := os.ReadDir(outputDir)
	if err != nil {
		return fmt.Errorf("無法讀取輸出資料夾: %v", err)
	}

	totalFiles := int64(len(files))
	var processedFiles int64

	// 獲取原始檔案名（不含副檔名）
	baseName := strings.TrimSuffix(filepath.Base(inputPath), filepath.Ext(inputPath))

	// 定義用於匹配日期的正則表達式
	dateRegex := regexp.MustCompile(`\d{4}[年/-]\d{1,2}[月/-]\d{1,2}[日]?`)
	// 定義用於匹配姓名的正則表達式
	nameRegex := regexp.MustCompile(`([\p{Han}]{2,3})(先生|小姐)`)

	var wg sync.WaitGroup
	sem := make(chan struct{}, 5) // 控制併發數量，防止資源耗盡
	var errors []string           // 儲存錯誤訊息
	var errorsMutex sync.Mutex    // 鎖，用於保護 errors 切片

	for _, file := range files {
		select {
		case <-cancelProcessing:
			statusLabel.SetText("已取消處理")
			return nil
		default:
			// 繼續處理
		}

		wg.Add(1)
		sem <- struct{}{}
		go func(file os.DirEntry) {
			defer wg.Done()
			defer func() { <-sem }()

			oldPath := filepath.Join(outputDir, file.Name())

			// 原始檔案直接略過
			if strings.TrimSuffix(file.Name(), filepath.Ext(oldPath)) == baseName {
				log.Println("原始檔案: ", strings.TrimSuffix(file.Name(), filepath.Ext(oldPath)), "跳過重新命名")
				return
			}

			// 提取 PDF 文字內容
			content, err := extractTextFromPDF(oldPath)
			if err != nil {
				log.Printf("無法提取檔案 %s 的文字: %v", oldPath, err)
				errorsMutex.Lock()
				errors = append(errors, fmt.Sprintf("檔案 %s：%v", file.Name(), err))
				errorsMutex.Unlock()
				return
			}

			// 從文字中提取日期
			date := extractDateFromText(content, dateRegex)
			if date == "" {
				log.Printf("未能從檔案 %s 中提取到日期", oldPath)
				errorsMutex.Lock()
				errors = append(errors, fmt.Sprintf("檔案 %s：未提取到日期", file.Name()))
				errorsMutex.Unlock()
				return
			}
			cleanedDate := cleanDateString(date)

			// 從文字中提取姓名
			name := extractNameFromText(content, nameRegex)
			if name == "" {
				log.Printf("未能從檔案 %s 中提取到姓名", oldPath)
				errorsMutex.Lock()
				errors = append(errors, fmt.Sprintf("檔案 %s：未提取到姓名", file.Name()))
				errorsMutex.Unlock()
				return
			}

			// 定義新的檔案名
			newName := fmt.Sprintf("%s-%s_%s.pdf", cleanedDate, name, baseName)
			newPath := filepath.Join(outputDir, newName)

			// 檢查是否已存在同名檔案，避免覆蓋
			if _, err := os.Stat(newPath); err == nil {
				log.Printf("檔案 %s 已存在，跳過重命名。", newPath)
				errorsMutex.Lock()
				errors = append(errors, fmt.Sprintf("檔案 %s：目標檔案已存在", file.Name()))
				errorsMutex.Unlock()
				return
			}

			// 重命名檔案
			err = os.Rename(oldPath, newPath)
			if err != nil {
				log.Printf("無法重命名檔案 %s: %v", oldPath, err)
				errorsMutex.Lock()
				errors = append(errors, fmt.Sprintf("檔案 %s：無法重命名", file.Name()))
				errorsMutex.Unlock()
				return
			}

			log.Printf("已將檔案 %s 重命名為 %s", oldPath, newName)

			// 更新進度條和狀態
			atomic.AddInt64(&processedFiles, 1)
			progress := float64(processedFiles) / float64(totalFiles)
			progressBar.SetValue(progress)
			statusLabel.SetText(fmt.Sprintf("正在處理：%d/%d", processedFiles, totalFiles))

		}(file)
	}

	wg.Wait()

	// 將錯誤訊息顯示給使用者
	if len(errors) > 0 {
		errorText := strings.Join(errors, "\n")
		errorList.SetText(errorText)
	}

	return nil
}

// 提取 PDF 文字內容
func extractTextFromPDF(filePath string) (string, error) {
	f, r, err := pdf.Open(filePath)
	if err != nil {
		return "", fmt.Errorf("無法打開檔案 %s: %v", filePath, err)
	}
	defer f.Close()

	// 假設每個檔案只有一頁
	page := r.Page(1)
	if page.V.IsNull() {
		return "", fmt.Errorf("無法獲取 PDF 檔案 %s 的第一頁", filePath)
	}

	content, err := page.GetPlainText(nil)
	if err != nil {
		return "", fmt.Errorf("無法提取檔案 %s 的文字: %v", filePath, err)
	}

	return content, nil
}

// 從文字中提取日期
func extractDateFromText(text string, dateRegex *regexp.Regexp) string {
	date := dateRegex.FindString(text)
	return date
}

// 從文字中提取姓名
func extractNameFromText(text string, nameRegex *regexp.Regexp) string {
	nameMatches := nameRegex.FindStringSubmatch(text)
	if len(nameMatches) >= 2 {
		return nameMatches[1]
	}
	return ""
}

// 清理日期字串，統一格式為 YYYY-MM-DD
func cleanDateString(date string) string {
	cleanedDate := strings.ReplaceAll(date, "/", "-")
	cleanedDate = strings.ReplaceAll(cleanedDate, "年", "-")
	cleanedDate = strings.ReplaceAll(cleanedDate, "月", "-")
	cleanedDate = strings.ReplaceAll(cleanedDate, "日", "")
	cleanedDate = strings.TrimSuffix(cleanedDate, "-")
	return cleanedDate
}

func deleteFile(path string) error {
	// 使用 os.Remove 刪除檔案
	err := os.Remove(path)
	if err != nil {
		return fmt.Errorf("無法刪除檔案 %s: %v", path, err)
	}
	return nil
}
