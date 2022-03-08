using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using ClosedXML.Excel;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace ExportSample.Service
{
    /// <summary>
    /// Export Service
    /// </summary>
    public static class ExportService
    {
        #region ===[ Export To Excel  ]============================================================

        /// <summary>
        /// Export to Excel and download file.
        /// </summary>
        public static void ExportExcelFile()
        {
            var records = DataService.GetData();
            var headers = new List<string[]> { new string[] { "First Name", "Last Name", "Email Address", "City" } };

            var wb = new XLWorkbook(); //create workbook
            var ws = wb.Worksheets.Add("Sample Export"); //add worksheet to workbook

            var rangeTitle = ws.Cell(1, 1).InsertData(headers); //insert headers to first row
            rangeTitle.AddToNamed("Headers");
            var titlesStyle = wb.Style;
            titlesStyle.Font.Bold = true;

            wb.NamedRanges.NamedRange("Headers").Ranges.Style = titlesStyle; //attach style to the range

            if (records.Any())
            {
                //insert data to from second row on
                ws.Cell(2, 1).InsertData(records);
                ws.Columns().AdjustToContents();
            }

            //save file to memory stream and return it as byte array
            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms);

                //Get current response
                var response = HttpContext.Current.Response;

                response.Clear();
                response.Buffer = true;
                response.AddHeader("content-disposition", "attachment; filename=Test_Data.xlsx");
                response.ContentType = "application/vnd.ms-excel";
                response.BinaryWrite(ms.ToArray());
                response.End();
            }
        }

        #endregion

        #region ===[ Export To CSV  ]==============================================================

        /// <summary>
        /// Export to CSV and download file.
        /// </summary>
        public static void ExportCsvFile()
        {
            //Get data you want to export
            //here you can use your own list or data table.
            var records = DataService.GetData();
            var sb = new StringBuilder();
            sb.AppendLine("First Name,Last Name,Email Address,City");
            foreach (var record in records)
            {
                sb.AppendLine($"{record.FirstName},{record.LastName},{record.Email},{record.City}");
            }

            //Get current response
            var response = HttpContext.Current.Response;
            response.BufferOutput = true;

            //clear response
            response.Clear();
            response.ClearHeaders();
            response.ContentEncoding = Encoding.Unicode;

            //provide file name.
            response.AddHeader("content-disposition", "attachment; filename=Test_Data.csv");
            response.ContentType = "application/vnd.ms-excel";
            response.Write(sb.ToString());
            response.End();
        }

        #endregion

        #region ===[ Export To Text  ]=============================================================

        /// <summary>
        /// Export to text and download file.
        /// </summary>
        public static void ExportTextFile()
        {
            var records = DataService.GetData();
            var sb = new StringBuilder();
            sb.AppendLine("First Name\tLast Name\tEmail Address\tCity");
            foreach (var record in records)
            {
                sb.AppendLine($"{record.FirstName}\t{record.LastName}\t{record.Email}\t{record.City}");
            }

            //Get current response
            var response = HttpContext.Current.Response;
            response.BufferOutput = true;

            //clear response
            response.Clear();
            response.ClearHeaders();
            response.ContentEncoding = Encoding.Unicode;

            //For csv export give file name as "Test_Data.csv"
            response.AddHeader("content-disposition", "attachment; filename=Test_Data.txt");
            response.ContentType = "text/plain";
            response.Write(sb.ToString());
            response.End();
        }

        #endregion

        #region ===[ Export To PDF  ]==============================================================

        /// <summary>
        /// Export to Pdf and download file.
        /// </summary>
        public static void ExportPdfFile()
        {
            //Create MemoryStream Object.
            var stream = new MemoryStream();

            //New Document with Page Size and Margins.
            var doc = new Document(PageSize.LETTER, 50f, 50f, 40f, 40f);

            // PdfWriter Object
            var writer = PdfWriter.GetInstance(doc, stream);
            writer.SetFullCompression();

            //Open document.
            doc.Open();

            //Font Factory to set font type, size and style.
            var style = FontFactory.GetFont(FontFactory.TIMES_ROMAN, 14f, Font.BOLD);

            //Add header.
            doc.Add(new Paragraph("Export PDF Sample", style)
            {
                SpacingAfter = 10,
                Alignment = 3
            });

            //New PdfPTable.
            var table = new PdfPTable(4) { WidthPercentage = 100 };

            // Set columns widths
            int[] widths = { 25, 25, 25, 25 };
            table.SetWidths(widths);

            style = FontFactory.GetFont(FontFactory.TIMES_ROMAN, 11f, Font.BOLD);

            table.AddCell(new PdfPCell(new Phrase("First Name", style)));
            table.AddCell(new PdfPCell(new Phrase("Last Name", style)));
            table.AddCell(new PdfPCell(new Phrase("Email Address", style)));
            table.AddCell(new PdfPCell(new Phrase("City", style)));

            style = FontFactory.GetFont(FontFactory.TIMES_ROMAN, 11f, Font.NORMAL);

            //Get Records.
            var records = DataService.GetData();

            foreach (var record in records)
            {
                table.AddCell(new PdfPCell(new Phrase(record.FirstName, style)));
                table.AddCell(new PdfPCell(new Phrase(record.LastName, style)));
                table.AddCell(new PdfPCell(new Phrase(record.Email, style)));
                table.AddCell(new PdfPCell(new Phrase(record.City, style)));
            }

            //Add table to document.
            doc.Add(table);

            //close document
            doc.Close();

            //convert stream to byte array.
            var byteArray = stream.ToArray();
            stream.Flush();

            //close stream
            stream.Close();

            if (byteArray != null)
            {
                //Write and download file.
                var response = HttpContext.Current.Response;
                response.Clear();
                response.AddHeader("Content-Disposition"
                    , "attachment; filename=Test_Data.pdf");
                response.AddHeader("Content-Length", byteArray.Length.ToString());
                response.ContentType = "application/pdf";
                response.BinaryWrite(byteArray);
                response.End();
            }
        }

        #endregion

        #region ===[ Export To HTML  ]=============================================================

        /// <summary>
        /// Export to Html and download file.
        /// </summary>
        public static void ExportHtmlFile()
        {
            const string headerFormat = "<th style='text-align: left;'>{0}</th>";

            var sb = new StringBuilder();

            sb.AppendFormat("<h1 style='font-size: 18px;font-weight: bold;'>{0}</h1>" , "Export HTML Sample");
            sb.Append("<table style='width:500px;'>");

            //Add Table Header
            sb.Append("<tr>");
            sb.AppendFormat(headerFormat, "First Name");
            sb.AppendFormat(headerFormat, "Last Name");
            sb.AppendFormat(headerFormat, "Email Address");
            sb.AppendFormat(headerFormat, "City");
            sb.Append("</tr>");

            var records = DataService.GetData();

            //Add Table Rows
            foreach (var record in records)
            {
                sb.Append($"<tr><td>{record.FirstName}</td><td>{record.LastName}</td><td>{record.Email}</td><td>{record.City}</td></tr>");
            }

            sb.Append("</table>");

            string htmlBody = sb.ToString();

            //Get current response
            var response = HttpContext.Current.Response;
            response.BufferOutput = true;

            //clear response
            response.Clear();
            response.ClearHeaders();

            //For HTML export give file name as "Test_Data.html"
            response.AddHeader("content-disposition", "attachment; filename=Test_Data.html");
            response.ContentType = "text/html";
            response.Write(htmlBody);
            response.End();
        }

        #endregion
    }
}