# PDFOperation Library

## Overview
PDFOperation is a utility library built on iTextSharp that provides functionality for generating, manipulating, and merging PDF documents. This library specializes in creating tabular PDF reports with proper pagination, headers, and formatting.

## License
This library utilizes iTextSharp, which is released under the GNU Affero General Public License (AGPL). When using this library, you must comply with the AGPL terms, including:

- If you modify this code, you must distribute your modifications under the AGPL.
- If you include this in an application accessed through a network, you must make the complete source code available to users of that application.

For more details, see: https://www.gnu.org/licenses/

## Features
- Generate PDF documents with table data
- Support for multiple pages with automatic pagination
- Custom headers and formatting
- Merge multiple PDF files
- Add page numbers to merged documents
- Support for Chinese character sets

## Usage

### Creating a PDF with Table Data
```csharp
// Create the headers and rows
List<string> headers = new List<string> { "Column 1", "Column 2", "Column 3" };
List<List<string>> rows = new List<List<string>> {
    new List<string> { "Data 1-1", "Data 1-2", "Data 1-3" },
    new List<string> { "Data 2-1", "Data 2-2", "Data 2-3" },
    new List<string> { "Data 3-1", "Data 3-2", "Data 3-3" }
};

// Call the GenPDF method (private, would need to be exposed or accessed through a public method)
string filePath = GenPDF(
    folderPath: @"C:\Output", 
    fileName: "Report.pdf", 
    tileName: "Sample Report", 
    headers: headers, 
    rows: rows
);
```

### Merging PDF Files
```csharp
List<string> pdfFiles = new List<string> {
    @"C:\Output\Report1.pdf",
    @"C:\Output\Report2.pdf"
};

PDFOperation.MergePdf(pdfFiles, @"C:\Output\MergedReport.pdf");
```

## Font Configuration
The library uses MingLiU (新細明體) by default, but you can specify a different font path.

---

# PDFOperation 函式庫

## 概述
PDFOperation 是基於 iTextSharp 構建的實用函式庫，提供生成、操作和合併 PDF 文檔的功能。該函式庫專門用於創建具有適當分頁、標題和格式的表格式 PDF 報告。

## 授權
此函式庫使用了 iTextSharp，該軟體在 GNU Affero 通用公共許可證 (AGPL) 下發布。當使用此函式庫時，您必須遵守 AGPL 條款，包括：

- 如果您修改此代碼，必須在 AGPL 下分發您的修改。
- 如果您將此函式庫納入通過網絡訪問的應用程序中，必須向該應用程序的用戶提供完整的源代碼。

更多詳情，請參見：https://www.gnu.org/licenses/

## 功能
- 生成帶有表格數據的 PDF 文檔
- 支持多頁自動分頁
- 自定義標題和格式
- 合併多個 PDF 文件
- 為合併的文件添加頁碼
- 支持中文字符集

## 使用方法

### 創建帶有表格數據的 PDF
```csharp
// 創建標題和行數據
List<string> headers = new List<string> { "欄位 1", "欄位 2", "欄位 3" };
List<List<string>> rows = new List<List<string>> {
    new List<string> { "數據 1-1", "數據 1-2", "數據 1-3" },
    new List<string> { "數據 2-1", "數據 2-2", "數據 2-3" },
    new List<string> { "數據 3-1", "數據 3-2", "數據 3-3" }
};

// 調用 GenPDF 方法（私有方法，需要通過公共方法訪問）
string filePath = GenPDF(
    folderPath: @"C:\Output", 
    fileName: "報告.pdf", 
    tileName: "範例報告", 
    headers: headers, 
    rows: rows
);
```

### 合併 PDF 文件
```csharp
List<string> pdfFiles = new List<string> {
    @"C:\Output\報告1.pdf",
    @"C:\Output\報告2.pdf"
};

PDFOperation.MergePdf(pdfFiles, @"C:\Output\合併報告.pdf");
```

## 字體配置
該函式庫默認使用新細明體，但您可以指定不同的字體路徑。