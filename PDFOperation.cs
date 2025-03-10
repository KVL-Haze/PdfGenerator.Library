/*
 * This program utilizes iTextSharp, a free software released under the GNU Affero General Public License.
 * iTextSharp Copyright (C) 1999-2023 by iText Group NV.
 * 
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Affero General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Affero General Public License for more details.
 *
 * You should have received a copy of the GNU Affero General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 *
 * AGPL licensing requires that if you modify this code or include it in an
 * application accessed through a network, you must make the complete source
 * code available to users of that application.
 */
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PdfGenerator.Library
{
    public class PDFOperation
    {
        /// <summary>
        /// 表格內容
        /// </summary>
        private class PDFTable
        {
            /// <summary>
            /// 標題
            /// </summary>
            public List<string> Headers { get; set; }
            /// <summary>
            /// 資料
            /// </summary>
            public List<List<string>> Rows { get; set; }
        }
        /// <summary>
        /// 產生含有標題與表格資料的PDF
        /// </summary>
        /// <param name="folderPath">資料夾路徑</param>
        /// <param name="fileName">檔案名稱</param>
        /// <param name="tileName">文本標題</param>
        /// <param name="headers">表格標題</param>
        /// <param name="rows">表格資料</param>
        /// <param name="fontPath">字體(預設新細明體)</param>
        /// <returns></returns>
        private string GenPDF(string folderPath, string fileName, string tileName, List<string> headers, List<List<string>> rows, string fontPath = @"C:\Windows\Fonts\mingliu.ttc,1")
        {
            if (!Directory.Exists(folderPath)) Directory.CreateDirectory(folderPath);
            string filePath = Path.Combine(folderPath, fileName);
            if (System.IO.File.Exists(filePath)) System.IO.File.Delete(filePath);

            PDFTable result = new PDFTable()
            {
                Headers = headers,
                Rows = rows
            };

            Document document = new Document(PageSize.A4.Rotate(), 30f, 30f, 15f, 15f);//橫向

            #region font
            BaseFont bfChinese = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

            iTextSharp.text.Font ChFont_small_8 = new iTextSharp.text.Font(bfChinese, 8f, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            iTextSharp.text.Font ChFont_small = new iTextSharp.text.Font(bfChinese, 9f, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            iTextSharp.text.Font ChFont_small_bold = new iTextSharp.text.Font(bfChinese, 9.5f, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
            iTextSharp.text.Font ChFont = new iTextSharp.text.Font(bfChinese, 11f, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            iTextSharp.text.BaseColor color_DarkBlue = new BaseColor(0, 51, 113);//(深藍色)
            iTextSharp.text.Font ChFont_color_DarkBlue = new iTextSharp.text.Font(bfChinese, 11f, iTextSharp.text.Font.NORMAL, color_DarkBlue);//(深藍色)
            iTextSharp.text.Font ChFont_bold = new iTextSharp.text.Font(bfChinese, 11.5f, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
            iTextSharp.text.Font ChFont_large = new iTextSharp.text.Font(bfChinese, 13f, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            iTextSharp.text.Font ChFont_large_bold = new iTextSharp.text.Font(bfChinese, 13.5f, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
            #endregion
            PdfWriter pdfWriter;
            using (var fileStream = new FileStream(filePath, FileMode.Create))
            {
                pdfWriter = PdfWriter.GetInstance(document, fileStream);
            }
            //開啟檔案
            document.Open();
            //無資料處理
            if (result.Rows[0].Count() == 0)
            {
                PdfPTable table = new PdfPTable(new float[] { 1f });
                PdfPCell cell;
                table = new PdfPTable(new float[] { 1f });
                cell = new PdfPCell(new Phrase(tileName, ChFont_large));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                cell.FixedHeight = 30;
                table.AddCell(cell);
                table.WidthPercentage = 100;
                document.Add(table);

                table = new PdfPTable(new float[] { 1f });
                cell = new PdfPCell(new Phrase("查無資料", ChFont_large));
                cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                cell.FixedHeight = 30;
                table.AddCell(cell);
                table.WidthPercentage = 100;
                document.Add(table);

                document.Close();
                return filePath;
            }

            //頁數
            int PageIndex = 1;
            int PageTotal = (result.Rows[0].Count() % 5) == 0 ? (result.Rows[0].Count() / 5) : (result.Rows[0].Count() / 5) + 1;

            for (int i = PageIndex; i <= PageTotal; i++)
            {
                //標題
                int rangeMin = i == 1 ? 0 : (i - 1) * 5;
                int rangeMax = (i) * 5 - 1;

                #region 列印資訊&頁次
                PdfPTable table_title_info = new PdfPTable(new float[] { 1f, 1f, 1f, 1f, 1f, 1f });
                PdfPCell cell_title_info;

                table_title_info = new PdfPTable(new float[] { 1f });
                cell_title_info = new PdfPCell(new Phrase(tileName, ChFont_large));
                cell_title_info.BorderWidth = 0;
                cell_title_info.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                cell_title_info.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                cell_title_info.FixedHeight = 30;
                table_title_info.AddCell(cell_title_info);
                table_title_info.WidthPercentage = 100;
                document.Add(table_title_info);

                table_title_info = new PdfPTable(new float[] { 1f });
                cell_title_info = new PdfPCell(new Phrase("頁數:" + i + " / " + PageTotal, ChFont_small));
                cell_title_info.BorderWidth = 0;
                cell_title_info.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                cell_title_info.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                cell_title_info.FixedHeight = 15;
                table_title_info.AddCell(cell_title_info);
                table_title_info.WidthPercentage = 80;
                document.Add(table_title_info);
                #endregion

                for (int rowIndex = 0; rowIndex < result.Rows.Count(); rowIndex++)
                {
                    List<string> items = result.Rows[rowIndex];
                    PdfPTable table_row = new PdfPTable(new float[] { 1.2f, 1f, 1f, 1f, 1f, 1f });
                    PdfPCell cell_row;

                    cell_row = new PdfPCell(new Phrase(result.Headers[rowIndex], ChFont_small));
                    cell_row.BackgroundColor = new BaseColor(212, 229, 238);
                    cell_row.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell_row.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell_row.FixedHeight = rowIndex == 0 ? 15 : 30;
                    if (rowIndex == result.Rows.Count() - 1)
                    {
                        cell_row.FixedHeight = 90;
                        if (result.Rows.Count() > 15)
                        {
                            cell_row.FixedHeight = 60;
                            if (result.Rows.Count() > 16)
                                cell_row.FixedHeight = 30;
                        }
                    }
                    table_row.AddCell(cell_row);
                    for (int cellIndex = rangeMin; cellIndex <= rangeMax; cellIndex++)
                    {
                        if (cellIndex < items.Count())
                        {
                            cell_row = new PdfPCell(new Phrase(items[cellIndex].Replace("\"", ""), ChFont_small));
                            if (rowIndex == 0) cell_row.BackgroundColor = new BaseColor(212, 229, 238);
                            cell_row.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                            cell_row.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                            cell_row.FixedHeight = rowIndex == 0 ? 15 : 30;
                            table_row.AddCell(cell_row);
                        }
                        else
                        {
                            cell_row = new PdfPCell(new Phrase(" ", ChFont_small));
                            if (rowIndex == 0) cell_row.BackgroundColor = new BaseColor(212, 229, 238);
                            cell_row.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                            cell_row.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                            cell_row.FixedHeight = rowIndex == 0 ? 15 : 30;
                            table_row.AddCell(cell_row);
                        }
                    }
                    table_row.WidthPercentage = 100;
                    document.Add(table_row);
                }

                if (i < PageTotal) document.NewPage();

            }
            document.Close();
            return filePath;
        }
        /// <summary>
        /// 合併PDF
        /// </summary>
        /// <param name="InFiles">合併來源(多個)</param>
        /// <param name="OutFile">合併結果</param>
        public static void MergePdf(List<string> InFiles, string OutFile)
        {
            string fileWithoutPageNo = OutFile.Replace(".pdf", "_tmp.pdf");

            using (FileStream stream = new FileStream(fileWithoutPageNo, FileMode.Create))
            using (Document doc = new Document())
            using (PdfCopy pdf = new PdfCopy(doc, stream))
            {
                doc.Open();

                PdfReader reader = null;
                PdfImportedPage page = null;

                //fixed typo
                InFiles.ForEach(file =>
                {
                    reader = new PdfReader(file);

                    for (int i = 0; i < reader.NumberOfPages; i++)
                    {
                        page = pdf.GetImportedPage(reader, i + 1);
                        pdf.AddPage(page);
                    }

                    pdf.FreeReader(reader);
                    reader.Close();

                    File.Delete(file);
                });
            }

            AddPdfPageNo(OutFile, fileWithoutPageNo);
        }
        /// <summary>
        /// 將PDF檔案加入頁碼後輸出
        /// </summary>
        /// <param name="outFile"></param>
        /// <param name="fileWithoutPageNo"></param>
        public static void AddPdfPageNo(string outFile, string fileWithoutPageNo)
        {
            using (var newPDF = new FileStream(outFile, FileMode.Create, FileAccess.ReadWrite))

            using (PdfReader pdfReader = new PdfReader(fileWithoutPageNo))
            using (PdfStamper pdfStamper = new PdfStamper(pdfReader, newPDF))
            {
                for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                {
                    PdfContentByte pdfContent = pdfStamper.GetOverContent(page);
                    iTextSharp.text.Rectangle mediabox = pdfReader.GetPageSize(page);

                    string pageText = $"{page}/{pdfReader.NumberOfPages}";

                    var offsetX = 37;
                    var offsetY = 51;

                    pdfContent.SetFontAndSize(BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, false), 10);
                    pdfContent.BeginText();
                    pdfContent.ShowTextAligned(Element.ALIGN_RIGHT, pageText, mediabox.Width - offsetX, offsetY, 0);
                    pdfContent.EndText();
                }

                pdfStamper.Close();
            }

            File.Delete(fileWithoutPageNo);
        }
    }
}
