using System;
using System.Activities;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;
using System.Activities.Statements;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Xml;




namespace Project_Graphic
{
    public class ExcelToLineGraph : CodeActivity
    {
        [RequiredArgument]
        [Category("Input")]
        public InArgument<string> ExcelFilePath { get; set; }

        [RequiredArgument]
        [Category("Input")]
        public InArgument<string> DataRange { get; set; }

        [RequiredArgument]
        [Category("Input")]
        public InArgument<double> MinValue { get; set; }

        [RequiredArgument]
        [Category("Input")]
        public InArgument<string> ChartTitle { get; set; }

        [Category("Output")]
        public OutArgument<string> OutputImagePath { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            string excelFilePath = ExcelFilePath.Get(context);
            string dataRange = DataRange.Get(context);
            string chartTitle = ChartTitle.Get(context);

            DataTable dataTable = new DataTable();

            // Excel uygulaması başlatılıyor
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                // Excel uygulamasını başlat elle başlatmam gerekti
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(excelFilePath);
                Excel.Worksheet worksheet = workbook.Sheets[1];
                Excel.Range range = worksheet.Range[dataRange];

                // Excel verisini object[,] türünde al
                object[,] cellValues = (object[,])range.Value;
                int rowCount = cellValues.GetLength(0);
                int columnCount = cellValues.GetLength(1);

                // DataTable'a sütunları ekle
                for (int col = 1; col <= columnCount; col++)
                {
                    dataTable.Columns.Add($"Column {col}", typeof(double));
                }

                // DataTable'a verileri ekle
                for (int row = 1; row <= rowCount; row++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int col = 1; col <= columnCount; col++)
                    {
                        dataRow[col - 1] = Convert.ToDouble(cellValues[row, col]);
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
                workbook = null;
                excelApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            // Grafik oluşturma
            Chart chart = new Chart();
            chart.Size = new System.Drawing.Size(600, 400);
            chart.ChartAreas.Add(new ChartArea("MainChartArea"));
            chart.Series.Add(new Series("DataSeries"));

            // Verileri grafik üzerine ekleme
            foreach (DataRow row in dataTable.Rows)
            {
                chart.Series["DataSeries"].Points.AddXY(row.Table.Columns[0].ColumnName, row[0]);
            }

            // Grafik başlığını ayarlama
            chart.Titles.Add(new Title(chartTitle));

            // Grafik resmini kaydetme
            string outputPath = $"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\\ExcelGraph.png";
            chart.SaveImage(outputPath, ChartImageFormat.Png);

            // Çıktı argümanını ayarlama
            OutputImagePath.Set(context, outputPath);
        }
    }
}
