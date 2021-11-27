using Contracter.Properties;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Contracter
{
	public partial class Form1 : Form
	{
		private const string ContractsWorksheetNameConfigName = "contracts_worksheet_name";
		private const string ContractsStartRowConfigName = "contracts_start_row";
		private const string ContractsCodeColumnConfigName = "contracts_code_column";
		private const string ContractsNameColumnConfigName = "contracts_name_column";
		private const string TemplateFileNameConfigName = "template_file_name";
		private const string TemplateWorksheetNameConfigName = "template_worksheet_name";
		private const string TemplateCellAddressConfigName = "template_cell_address";
		private const string DoubleFileSuffixConfigName = "double_file_suffix";

		private static readonly string InvalidPattern = $"[{Regex.Escape(new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars()))}]";

		private static string ContractsWorksheetName => ConfigurationManager.AppSettings[ContractsWorksheetNameConfigName];
		private static int ContractsStartRow => int.Parse(ConfigurationManager.AppSettings[ContractsStartRowConfigName]);
		private static int ContractsCodeColumn => int.Parse(ConfigurationManager.AppSettings[ContractsCodeColumnConfigName]);
		private static int ContractsNameColumn => int.Parse(ConfigurationManager.AppSettings[ContractsNameColumnConfigName]);
		private static string TemplateFileName => ConfigurationManager.AppSettings[TemplateFileNameConfigName];
		private static string TemplateWorksheetName => ConfigurationManager.AppSettings[TemplateWorksheetNameConfigName];
		private static string TemplateCellAddress => ConfigurationManager.AppSettings[TemplateCellAddressConfigName];
		private static string DoubleFileSuffix => ConfigurationManager.AppSettings[DoubleFileSuffixConfigName];

		private string sourceFile;
		private string destinationDirectory;

		private static string StripInvalidPathChars(string path) =>
			Regex.Replace(path, InvalidPattern, "_", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase | RegexOptions.Singleline);

		private static string EnsureUniqueFileName(string path)
		{
			var index = 1;
			while (File.Exists(path))
			{
				var name = Path.GetFileNameWithoutExtension(path);
				var uniqueIndex = string.Format(DoubleFileSuffix, index);

				path = Path.Combine(Path.GetDirectoryName(path), name + uniqueIndex + Path.GetExtension(path));
				++index;
			}

			return path;
		}

		private static string EnsureUniqueDirectoryName(string path)
		{
			var index = 1;
			while (Directory.Exists(path))
			{
				var name = Path.GetFileName(path);
				var uniqueIndex = string.Format(DoubleFileSuffix, index);

				path = Path.Combine(Path.GetDirectoryName(path), name + uniqueIndex);
				++index;
			}

			return path;
		}

		public Form1()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			var result = openFileDialog1.ShowDialog();
			if (result == DialogResult.OK)
			{
				textBox1.Text = openFileDialog1.FileName;
			}
		}

		private void button2_Click(object sender, EventArgs e)
		{
			var result = folderBrowserDialog1.ShowDialog();
			if (result == DialogResult.OK)
			{
				textBox2.Text = folderBrowserDialog1.SelectedPath;
			}
		}

		private void button3_Click(object sender, EventArgs e)
		{
			if (!backgroundWorker1.IsBusy)
			{
				if (!File.Exists(textBox1.Text))
				{
					MessageBox.Show("Некорректный путь к файлу с отчётом для выборки!", "Ошибка исходного файла", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				else
				{
					if (!Directory.Exists(textBox2.Text))
					{
						MessageBox.Show("Некорректный путь к папке назначения!", "Ошибка папки назначения", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					else
					{
						button3.Text = "Отменить";

						sourceFile = textBox1.Text;
						destinationDirectory = textBox2.Text;

						backgroundWorker1.RunWorkerAsync();
					}
				}
			}
			else
			{
				backgroundWorker1.CancelAsync();
			}
		}

		private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
		{
			var worker = sender as BackgroundWorker;
			try
			{
				using (var contractsStream = File.OpenRead(sourceFile))
				{
					using (var contractsPackage = new ExcelPackage(contractsStream))
					{
						var contractsWorksheet = contractsPackage.Workbook.Worksheets[ContractsWorksheetName];
						var contractsRows = contractsWorksheet.Dimension.End.Row - 1;
						for (var i = ContractsStartRow; i <= contractsRows && !worker.CancellationPending; i++)
						{
							var contractCode = contractsWorksheet.Cells[i, ContractsCodeColumn].Text;
							var contractName = contractsWorksheet.Cells[i, ContractsNameColumn].Text;

							var outputFolder = EnsureUniqueDirectoryName(Path.Combine(destinationDirectory, StripInvalidPathChars(contractCode)));
							Directory.CreateDirectory(outputFolder);

							var outputFile = Path.Combine(outputFolder, TemplateFileName);

							using (var templateStream = File.OpenRead(TemplateFileName))
							{
								using (var templatePackage = new ExcelPackage(templateStream))
								{
									var templateWorksheet = templatePackage.Workbook.Worksheets[TemplateWorksheetName];

									var templateText = templateWorksheet.Cells[TemplateCellAddress].Text;
									templateText = string.Format(templateText, contractName, contractCode);

									templateWorksheet.Cells[TemplateCellAddress].Value = templateText;

									using (var outputStream = File.Create(outputFile))
									{
										templatePackage.SaveAs(outputStream);
									}
								}
							}

							worker.ReportProgress((int)((decimal)(i + 1) / (decimal)(contractsRows - ContractsStartRow) * 100.0M));
						}
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "Ошибка обработки", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
		{
			progressBar1.Value = e.ProgressPercentage;
		}

		private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			MessageBox.Show("Генерация завершена", "Процесс завершён", MessageBoxButtons.OK, MessageBoxIcon.Information);
			button3.Text = "Начать";
			progressBar1.Value = 0;
		}
	}
}
