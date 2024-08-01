using Parquet.Schema;
using Parquet;
using System.Data;
using OxyPlot;
using OxyPlot.Series;
using OxyPlot.WindowsForms;
using DataColumn = System.Data.DataColumn;
using OxyPlot.Axes;
using System.Text;
using OfficeOpenXml;

namespace ArcVera_Tech_Test
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private async void btnImportEra5_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Parquet files (*.parquet)|*.parquet|All files (*.*)|*.*";
                openFileDialog.Title = "Select a Parquet File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string parquetFilePath = openFileDialog.FileName;
                    DataTable era5DataTable = await ReadParquetFileAsync(parquetFilePath);
                    dgImportedEra5.DataSource = era5DataTable;
                    PlotU10DailyValues(era5DataTable);
                }
            }
        }

        private async Task<DataTable> ReadParquetFileAsync(string parquetFilePath)
        {
            using (Stream fileStream = File.OpenRead(parquetFilePath))
            {
                using (var parquetReader = await ParquetReader.CreateAsync(fileStream))
                {
                    DataTable era5DataTable = new DataTable();

                    for (int i = 0; i < parquetReader.RowGroupCount; i++)
                    {
                        using (ParquetRowGroupReader groupReader = parquetReader.OpenRowGroupReader(i))
                        {
                            // Create columns
                            foreach (DataField field in parquetReader.Schema.GetDataFields())
                            {
                                if (!era5DataTable.Columns.Contains(field.Name))
                                {
                                    Type columnType = field.HasNulls ? typeof(object) : field.ClrType;
                                    era5DataTable.Columns.Add(field.Name, columnType);
                                }

                                // Read values from Parquet column
                                DataColumn column = era5DataTable.Columns[field.Name];
                                Array values = (await groupReader.ReadColumnAsync(field)).Data;
                                for (int j = 0; j < values.Length; j++)
                                {
                                    if (era5DataTable.Rows.Count <= j)
                                    {
                                        era5DataTable.Rows.Add(era5DataTable.NewRow());
                                    }
                                    era5DataTable.Rows[j][field.Name] = values.GetValue(j);
                                }
                            }
                        }
                    }

                    return era5DataTable;
                }
            }
        }

        private void PlotU10DailyValues(DataTable era5DataTable)
        {
            var u10PlotModel = new PlotModel { Title = "Daily u10 Values" };
            var u10LineSeries = new LineSeries { Title = "u10" };

            var dailyGroupedData = era5DataTable.AsEnumerable()
                .GroupBy(row => DateTime.Parse(row["date"].ToString()))
                .Select(g => new
                {
                    Date = g.Key,
                    U10Average = g.Average(row => Convert.ToDouble(row["u10"]))
                })
                .OrderBy(data => data.Date);

            foreach (var data in dailyGroupedData)
            {
                u10LineSeries.Points.Add(new DataPoint(DateTimeAxis.ToDouble(data.Date), data.U10Average));
            }

            u10PlotModel.Series.Add(u10LineSeries);
            plotView1.Model = u10PlotModel;
        }

        private void btnExportCsv_Click(object sender, EventArgs e)
        {
            // Complete here
            DataTable era5DataTable = (DataTable)dgImportedEra5.DataSource;
            if (era5DataTable == null)
            {
                MessageBox.Show("No data to export", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                saveFileDialog.Title = "Save as CSV";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string csvFilePath = saveFileDialog.FileName;
                    ExportDataTableToCsv(era5DataTable, csvFilePath);
                }
            }
        }

        private void ExportDataTableToCsv(DataTable era5DataTable, string csvFilePath)
        {
            var start = DateTime.Now;
            var columnNames = era5DataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();
            var csvBuilder = new StringBuilder();
            csvBuilder.AppendLine(string.Join(",", columnNames));
            foreach (DataRow row in era5DataTable.Rows)
            {
                csvBuilder.AppendLine(string.Join(",", row.ItemArray));
            }
            File.WriteAllText(csvFilePath, csvBuilder.ToString());
            var end = DateTime.Now;
            MessageBox.Show($"Exported {era5DataTable.Rows.Count} rows to CSV in {(end - start).TotalSeconds} seconds", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            // Complete here
            DataTable era5DataTable = (DataTable)dgImportedEra5.DataSource;
            if (era5DataTable == null)
            {
                MessageBox.Show("No data to export", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "XLSX files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveFileDialog.Title = "Save as XLSX";


                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string excelFilePath = saveFileDialog.FileName;
                    ExportDataTableToExcel(era5DataTable, excelFilePath);
                }
            }
        }


        private void ExportDataTableToExcel(DataTable era5DataTable, string excelFilePath)
        {
            var start = DateTime.Now;
            int rowsPerPage = 1000000;
            int rowsNotImportedYet = era5DataTable.Rows.Count;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using ExcelPackage package = new ExcelPackage();
            for (int iteration = 0; rowsNotImportedYet > 0; iteration++)
            {
                using DataTable dt = new DataTable();
                for (int i = 0; i < era5DataTable.Columns.Count; i++)
                {
                    dt.Columns.Add(era5DataTable.Columns[i].ColumnName, era5DataTable.Columns[i].DataType);
                }


                int rowsAlreadyImported = era5DataTable.Rows.Count - rowsNotImportedYet;
                int rowsToImport = Math.Min(rowsPerPage, rowsNotImportedYet);
                for (int i = rowsAlreadyImported; i < rowsAlreadyImported + rowsToImport; i++)
                {
                    dt.ImportRow(era5DataTable.Rows[i]);
                }
                rowsNotImportedYet -= rowsToImport;


                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"Sheet{iteration}");
                var filledRange = worksheet.Cells["A1"].LoadFromDataTable(dt, true);
                int rowCount = dt.Rows.Count;
                var cf = worksheet.ConditionalFormatting.AddExpression(worksheet.Cells[$"A1:E{rowCount + 1}"]);
                cf.Formula = "IF($E1<0,1,0)";
                cf.Style.Fill.BackgroundColor.SetColor(Color.Red);
                cf.Style.Font.Color.SetColor(Color.White);


                dt.Clear();
            }
            package.SaveAs(new FileInfo(excelFilePath));


            var end = DateTime.Now;
            MessageBox.Show($"Exported {era5DataTable.Rows.Count} rows to Excel in {(end - start).TotalSeconds} seconds", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }





    }
}
