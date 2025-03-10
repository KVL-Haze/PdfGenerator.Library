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
The library uses MingLiU (�s�ө���) by default, but you can specify a different font path.

---

# PDFOperation �禡�w

## ���z
PDFOperation �O��� iTextSharp �c�ت���Ψ禡�w�A���ѥͦ��B�ާ@�M�X�� PDF ���ɪ��\��C�Ө禡�w�M���Ω�Ыب㦳�A������B���D�M�榡����榡 PDF ���i�C

## ���v
���禡�w�ϥΤF iTextSharp�A�ӳn��b GNU Affero �q�Τ��@�\�i�� (AGPL) �U�o���C��ϥΦ��禡�w�ɡA�z������u AGPL ���ڡA�]�A�G

- �p�G�z�ק惡�N�X�A�����b AGPL �U���o�z���ק�C
- �p�G�z�N���禡�w�ǤJ�q�L�����X�ݪ����ε{�Ǥ��A�����V�����ε{�Ǫ��Τᴣ�ѧ��㪺���N�X�C

��h�Ա��A�аѨ��Ghttps://www.gnu.org/licenses/

## �\��
- �ͦ��a�����ƾڪ� PDF ����
- ����h���۰ʤ���
- �۩w�q���D�M�榡
- �X�֦h�� PDF ���
- ���X�֪����K�[���X
- �������r�Ŷ�

## �ϥΤ�k

### �Ыرa�����ƾڪ� PDF
```csharp
// �Ыؼ��D�M��ƾ�
List<string> headers = new List<string> { "��� 1", "��� 2", "��� 3" };
List<List<string>> rows = new List<List<string>> {
    new List<string> { "�ƾ� 1-1", "�ƾ� 1-2", "�ƾ� 1-3" },
    new List<string> { "�ƾ� 2-1", "�ƾ� 2-2", "�ƾ� 2-3" },
    new List<string> { "�ƾ� 3-1", "�ƾ� 3-2", "�ƾ� 3-3" }
};

// �ե� GenPDF ��k�]�p����k�A�ݭn�q�L���@��k�X�ݡ^
string filePath = GenPDF(
    folderPath: @"C:\Output", 
    fileName: "���i.pdf", 
    tileName: "�d�ҳ��i", 
    headers: headers, 
    rows: rows
);
```

### �X�� PDF ���
```csharp
List<string> pdfFiles = new List<string> {
    @"C:\Output\���i1.pdf",
    @"C:\Output\���i2.pdf"
};

PDFOperation.MergePdf(pdfFiles, @"C:\Output\�X�ֳ��i.pdf");
```

## �r��t�m
�Ө禡�w�q�{�ϥηs�ө���A���z�i�H���w���P���r����|�C