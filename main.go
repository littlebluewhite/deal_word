package main

import (
	"fmt"
	"github.com/unidoc/unioffice/document"
	"os"
)

//TIP <p>To run your code, right-click the code and select <b>Run</b>.</p> <p>Alternatively, click
// the <icon src="AllIcons.Actions.Execute"/> icon in the gutter and select the <b>Run</b> menu item from here.</p>

func main() {
	// 打開要處理的 Word 文檔
	doc, err := document.Open("000.docx")
	if err != nil {
		fmt.Println("Error opening document:", err)
		return
	}

	// 初始化一個變量來保存當前頁數
	pageCount := 1

	// 創建一個新文檔來保存每一頁的內容
	newDoc := document.New()

	// 遍歷段落來模擬頁分割
	for _, para := range doc.Paragraphs() {
		// 添加當前段落到新文檔
		newPara := newDoc.AddParagraph()
		for _, run := range para.Runs() {
			newRun := newPara.AddRun()
			newRun.AddText(run.Text())
		}

		// 假設在特定條件下進行頁分割，這裡可以根據具體需求調整
		// 比如，每 X 段落分割成一個新頁，或根據某些標記進行分頁。
		if conditionToSplitPage() { // 這裡需要你自己定義頁分割的條件
			// 保存當前文檔
			fileName := fmt.Sprintf("output_page_%d.docx", pageCount)
			saveDoc(newDoc, fileName)
			pageCount++

			// 創建新文檔來保存下一頁的內容
			newDoc = document.New()
		}
	}

	// 保存最後一頁
	fileName := fmt.Sprintf("output_page_%d.docx", pageCount)
	saveDoc(newDoc, fileName)
}

// saveDoc 用來保存文檔
func saveDoc(doc *document.Document, fileName string) {
	out, err := os.Create(fileName)
	if err != nil {
		fmt.Println("Error creating file:", err)
		return
	}
	defer out.Close()

	doc.Save(out)
}

// conditionToSplitPage 是分頁的邏輯，可以根據實際需求進行調整
func conditionToSplitPage() bool {
	// 可以根據特定條件進行頁分割
	// 這裡僅作為示例，可以改為段落數、特定文本、或其他標記來分割頁面
	return false

}
