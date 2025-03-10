<div align="center">
    <h1>
    PDFOperation Library
    </h1>
    <p>
    <b>簡易的PDF生成與處理工具，基於iTextSharp構建</b><br>
    <em>Efficient PDF generation and processing tool built on iTextSharp</em>
    </p>
    <a href="https://www.gnu.org/licenses/agpl-3.0"><img src="https://img.shields.io/badge/License-AGPL%203.0-blue.svg" alt="license"></a>
    <a href="#"><img src="https://img.shields.io/badge/Platform-Windows-lightgrey" alt="platform"></a>
    <a href="#"><img src="https://img.shields.io/badge/.NET-Framework%204.5+-orange" alt="framework"></a>
</div>

## Overview

PDFOperation is a utility library built on iTextSharp that provides functionality for generating, manipulating, and merging PDF documents. This library specializes in creating tabular PDF reports with proper pagination, headers, and formatting.

## Key Features

- **Generate PDF documents with table data** - Create professional reports with tabular data
- **Multiple pages with automatic pagination** - Handle large datasets across multiple pages
- **Custom headers and formatting** - Customize the appearance of your PDF documents
- **Merge multiple PDF files** - Combine separate PDFs into a single document
- **Automatic page numbering** - Add page numbers to merged documents
- **Support for Chinese character sets** - Full Unicode support for international text

---

<table align="center">
  <tr>
    <td align="center"><b>PDF Generation</b><br>Create well-formatted PDF reports from data tables</td>
    <td align="center"><b>PDF Merging</b><br>Combine multiple PDFs into a single document</td>
  </tr>
</table>

## Install

To use PDFOperation in your project, install via NuGet or add a reference to the library in your solution.

```csharp
// Install via NuGet Package Manager Console
Install-Package iTextSharp
// Then reference PDFOperation in your project
```

## Usage

### Creating a PDF with Table Data

```csharp
// Create the headers and rows
List<string> headers = new List<string> { "Column 1", "Column 2", "Column 3" };
List<List<string>> rows = new List<List<string>> {
    new List<string> { "Data 1-1", "Data 1-2", "Data 1-3" },
    new List<string> { "Data 2-1", "Data 2-2", "Data 2-3" }
};

// Call the genPDF method (private, would need to be exposed or accessed through a public method)
string filePath = genPDF(
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

The library uses MingLiU (新細明體) by default, but you can specify a different font path:

```csharp
// Specify custom font path
string filePath = genPDF(
    // ... other parameters
    fontPath: @"C:\Windows\Fonts\kaiu.ttf" // 指定標楷體
);
```

## License

This library utilizes iTextSharp, which is released under the GNU Affero General Public License (AGPL). When using this library, you must comply with the AGPL terms, including:

- If you modify this code, you must distribute your modifications under the AGPL.
- If you include this in an application accessed through a network, you must make the complete source code available to users of that application.

For more details, see: [GNU Licenses](https://www.gnu.org/licenses/)

---

## 中文說明

<div align="center">
    <h2>
    PDFOperation 函式庫
    </h2>
    <p>
    <b>簡易的PDF生成與處理工具，基於iTextSharp構建</b>
    </p>
</div>

### 概述

PDFOperation 是基於 iTextSharp 構建的實用函式庫，提供生成、操作和合併 PDF 文檔的功能。該函式庫專門用於創建具有適當分頁、標題和格式的表格式 PDF 報告。

### 主要功能

- **生成帶有表格數據的 PDF 文檔** - 創建專業的表格式報告
- **多頁自動分頁** - 處理跨多頁的大型數據集
- **自定義標題和格式** - 自定義 PDF 文檔的外觀
- **合併多個 PDF 文件** - 將多個PDF文件合併為一個文檔
- **自動頁碼添加** - 為合併後的文檔添加頁碼
- **支持中文字符集** - 完整支持中文字符

### 使用方法

#### 創建帶有表格數據的 PDF

```csharp
// 創建標題和行數據
List<string> headers = new List<string> { "欄位 1", "欄位 2", "欄位 3" };
List<List<string>> rows = new List<List<string>> {
    new List<string> { "數據 1-1", "數據 1-2", "數據 1-3" },
    new List<string> { "數據 2-1", "數據 2-2", "數據 2-3" }
};

// 調用 genPDF 方法（私有方法，需要通過公共方法訪問）
string filePath = genPDF(
    folderPath: @"C:\Output", 
    fileName: "報告.pdf", 
    tileName: "範例報告", 
    headers: headers, 
    rows: rows
);
```

#### 合併 PDF 文件

```csharp
List<string> pdfFiles = new List<string> {
    @"C:\Output\報告1.pdf",
    @"C:\Output\報告2.pdf"
};

PDFOperation.MergePdf(pdfFiles, @"C:\Output\合併報告.pdf");
```

### 授權說明

此函式庫使用了 iTextSharp，該軟體在 GNU Affero 通用公共許可證 (AGPL) 下發布。當使用此函式庫時，您必須遵守 AGPL 條款，包括：

- 如果您修改此代碼，必須在 AGPL 下分發您的修改。
- 如果您將此函式庫納入通過網絡訪問的應用程序中，必須向該應用程序的用戶提供完整的源代碼。

更多詳情，請參見：[GNU 許可證](https://www.gnu.org/licenses/)
