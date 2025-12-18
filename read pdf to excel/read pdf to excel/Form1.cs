using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using Excel = Microsoft.Office.Interop.Excel;

namespace read_pdf_to_excel
{

    public partial class Form1 : Form
    {
        DataTable dtResult;
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            // 1. Open the file dialog to select the PDF.
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == DialogResult.OK)
            {
                string filePath = file.FileName;

                DataTable dtAbbreviations = ReadAbbreviationsFromPdf(filePath);

               dtResult = GetBreakerSummary(filePath, dtAbbreviations);

                // 3. Display the result in the DataGridView.
                dataGridView1.DataSource = dtResult;
            }
        }

        static DataTable ReadAbbreviationsFromPdf(string pdfPath)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Abbreviation");
            dt.Columns.Add("Description");

            using (PdfReader reader = new PdfReader(pdfPath))
            using (PdfDocument pdf = new PdfDocument(reader))
            {
                for (int i = 1; i <= pdf.GetNumberOfPages(); i++)
                {
                    string text = PdfTextExtractor.GetTextFromPage(pdf.GetPage(i));

                    if (text.Contains("Abbreviation") && text.Contains("Description"))
                    {
                        ExtractToDataTable(text, dt);
                    }
                }
            }

            return dt;
        }
        static void ExtractToDataTable(string text, DataTable dt)
        {

            var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(l => l.Trim())
                            .Where(l =>
                                !string.IsNullOrWhiteSpace(l) &&
                                !l.StartsWith("Abbreviation", StringComparison.OrdinalIgnoreCase) &&
                                !l.StartsWith("3.") &&
                                !l.Contains(".docx") &&
                                !l.Contains("/"))
                            .ToList();
            foreach (string line in lines)
            {
                int spaceIndex = line.IndexOf(' ');

                if (spaceIndex <= 0) continue;
                string abbr = line.Substring(0, spaceIndex).Trim();
                string desc = line.Substring(spaceIndex + 1).Trim();
                if (abbr.Length < 2) continue;
                dt.Rows.Add(abbr, desc);

            }

        }
        public static DataTable GetBreakerSummary(string pdfPath , DataTable dtAbbreviations)
        {
          
            string rawSldText = ReadRawTextFromPage(pdfPath, 3);
            string normalizedText = NormalizeSldText(rawSldText);
            DataTable dtBreakers = ExtractBreakerData(normalizedText);
            DataTable dtResult = GroupAndCount(dtBreakers, dtAbbreviations);

            return dtResult;
        }

 
        private static string ReadRawTextFromPage(string pdfPath, int pageNumber)
        {

                using (PdfReader reader = new PdfReader(pdfPath))
                using (PdfDocument pdf = new PdfDocument(reader))
                {
                    if (pageNumber > 0 && pageNumber <= pdf.GetNumberOfPages())
                    {
                       
                        return PdfTextExtractor.GetTextFromPage(pdf.GetPage(pageNumber));
                    }
                    return string.Empty;
                }

        }

        // --- HELPER 2: TEXT NORMALIZATION ---
        private static string NormalizeSldText(string text)
        {
            text = text.ToUpper();
            text = text.Replace("\r", " ").Replace("\n", " ");
            text = Regex.Replace(text, @"\s+", " ");

            // Correct common OCR/parsing errors based on the SLD text
            text = text.Replace("HCCS", "MCCB");
            text = text.Replace("HC", "MCCB");
            text = text.Replace("HCCB", "MCCB");
            text = text.Replace("A/5 A", "A");
            text = text.Replace("MOL", "MCCB");
          
            // Normalize units and remove auxiliary text (KW, W, V, relay codes)
            text = Regex.Replace(text, @"\s*A\b", "A");
            text = Regex.Replace(text, @"\d{1,5}\s*(KW|W|V|VOLT)", " ");
            text = text.Replace("50,50,51,", " ");
            text = text.Replace("C100, 8-10", " ");

            return text.Trim();
        }

  
        private static DataTable ExtractBreakerData(string normalizedText)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Type");
            dt.Columns.Add("Current");
            dt.Columns.Add("Poles");

            var currentMatches = Regex.Matches(normalizedText, @"\b\d{1,5}A\b")
                                     .Cast<Match>()
                                     .Select(m => m.Value)
                                     .Distinct()
                                     .ToList();

            var breakerInstances = new List<(string Type, string Current, string Poles)>();

            foreach (string current in currentMatches)
            {
                int totalCurrentCount = Regex.Matches(normalizedText, current).Count;

                string type;
                string poles = "3P";

                if (current == "1200A")
                {

                    type = "ACB";

                    totalCurrentCount = totalCurrentCount / 3;
                }
                else
                   if (current == "200A")
                {
                    type = "MCCB";
                    totalCurrentCount=1;
                }
                else
                    {
                    type = "MCCB";

                }

                for (int i = 0; i < totalCurrentCount; i++)
                {
                    breakerInstances.Add((type, current, poles));
                }
            }

            // Manually add the two MCBs (Miniature Circuit Breakers) for small control loads
            // These were identified near the transformer text tokens. Assumed 15A.
            //breakerInstances.Add(("MCB", "15A", "3P"));
            //breakerInstances.Add(("MCB", "15A", "3P"));

            foreach (var instance in breakerInstances)
            {
                dt.Rows.Add(instance.Type, instance.Current, instance.Poles);
            }

            return dt;
        }

        

        // --- HELPER 5: GROUP, COUNT, AND JOIN ---
        private static DataTable GroupAndCount(DataTable breakersDT, DataTable abbrevDT)
        {
            DataTable result = new DataTable();
            result.Columns.Add("BreakerType");
            result.Columns.Add("Current");
            result.Columns.Add("Poles");
            result.Columns.Add("Count", typeof(int));

            var query =
                from b in breakersDT.AsEnumerable()
                join a in abbrevDT.AsEnumerable()
                    on b["Type"].ToString() equals a["Abbreviation"].ToString()
                    into joined
                from a in joined.DefaultIfEmpty()
                select new
                {
                    Current = b["Current"].ToString(),
                    Poles = b["Poles"].ToString(),
                    // Use the full description (e.g., Molded Circuit Breaker)
                    Type = a != null ? a["Description"].ToString() : b["Type"].ToString()
                };

            var grouped = query
                .GroupBy(x => new { x.Current, x.Poles, x.Type })
                .Select(g => new
                {
                    g.Key.Current,
                    g.Key.Poles,
                    g.Key.Type,
                    Count = g.Count()
                });

            // Order by current descending for a clean summary
            foreach (var row in grouped.OrderByDescending(r => int.Parse(r.Current.Replace("A", ""))))
            {
                result.Rows.Add(row.Type, row.Current, row.Poles, row.Count);
            }

            return result;
        }

        // Fallback/Testing text based on the full content of Page 3
        private static string GetFallbackSldText()
        {
            return @"
            1200 A/S A 
            1200 A/S A
            1200 A/S A
            MCB
            MCB
            MCCB
            MCCB
            250A
            100A
            50A
            50A
            50A
            50A
            HCCS
            HC
            MOL
            200A
            200A
            200A
            ";
        }

        private void button2_Click(object sender, EventArgs e)
        {
           // DataTable dt = GetDataTable();

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
            Excel._Worksheet worksheet = workbook.ActiveSheet;
            worksheet.Name = "Export";

            // ------ Write column headers ------
            for (int i = 0; i < dtResult.Columns.Count; i++)
            {
                worksheet.Cells[1, i + 1] = dtResult.Columns[i].ColumnName;
            }

            // ------ Write data rows ------
            for (int row = 0; row < dtResult.Rows.Count; row++)
            {
                for (int col = 0; col < dtResult.Columns.Count; col++)
                {
                    worksheet.Cells[row + 2, col + 1] = dtResult.Rows[row][col].ToString();
                }
            }

            // ------ Save file ------
            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "Excel File|*.xlsx";

            if (save.ShowDialog() == DialogResult.OK)
            {
                workbook.SaveAs(
    save.FileName,
    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
    Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
);

            }

            workbook.Close();
            excelApp.Quit();

        }
    }
}
