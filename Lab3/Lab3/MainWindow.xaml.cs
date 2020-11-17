using System;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace Lab3
{
	public partial class MainWindow : Window
	{
		private readonly string _path = @"D:\5_sem\TMP\3\Lab3\VBA\Lab3.1.xlsm";

		private Excel.Application _excel;
		private Excel.Workbook _workBook;

		public MainWindow() => InitializeComponent();

		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			try
			{
				Excel.Application excel = new Excel.Application { Visible = true };
				Excel.Workbook workBook = excel.Workbooks.Open(_path);
				Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Sheets[1];

				int funcsCount = (int)workSheet.Cells[1, "F"].Value;
				for (int i = 1; i <= funcsCount; i++)
					ComboBoxFuncs.Items.Add(workSheet.Cells[i, "A"].Value);

				_excel = excel;
				_workBook = workBook;
			}
			catch (Exception er)
			{
				MessageBox.Show(er.Message);
				Close();
			}
		}

		private void ComboBoxFuncs_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
		{
			Excel.Worksheet workSheet = (Excel.Worksheet)_workBook.Sheets[2];
			workSheet.Cells[2, "F"].Value = (sender as System.Windows.Controls.ComboBox).SelectedIndex + 1;
		}

		private void ButtonApply_Click(object sender, RoutedEventArgs e)
		{
			Excel.Worksheet workSheet2 = (Excel.Worksheet)_workBook.Sheets[2];
			Excel.Worksheet workSheet3 = (Excel.Worksheet)_workBook.Sheets[3];

			int rowIndex = 0;

			for (int x = 0; x <= 10; x++)
			{
				workSheet2.Cells[3, "F"].Value = x;
				double y = workSheet2.Cells[6, "F"].Value;

				workSheet3.Cells[rowIndex + 1, "H"].Value = x;
				workSheet3.Cells[rowIndex + 1, "I"].Value = y;

				rowIndex++;
			}

			Excel.ChartObjects chartObjs = (Excel.ChartObjects)workSheet3.ChartObjects(Type.Missing);
			Excel.ChartObject myChart = chartObjs.Add(20, 60, 200, 200);
			Excel.Chart chart = myChart.Chart;
			Excel.SeriesCollection seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection(Type.Missing);
			Excel.Series series = seriesCollection.NewSeries();
			series.XValues = workSheet3.get_Range("H1", "H10");
			series.Values = workSheet3.get_Range("I1", "I10");
			chart.ChartType = Excel.XlChartType.xlXYScatterSmooth;

		}
	}
}