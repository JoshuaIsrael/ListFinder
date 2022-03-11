using ListFinder.Libraries;
using ListFinder.Models;
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace ListFinder
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<string> keysText = new List<string>();
        private List<string> resultText = new List<string>();
        private ExcelModel excelModel = new ExcelModel();

        private List<Control> controls = new List<Control>();

        public MainWindow()
        {
            InitializeComponent();

            controls.AddRange(new List<Control>
            {
                btnOpenFile,
                txtStartingRowNumber,
                txtColumnNumber,
                btnExtractExcelFile,
                txtSearchDirectory,
                btnBrowse,
                btnStartSearch,
                txtPrefix,
                txtSuffix,
                txtFileExtensions
        });
        }

        private async void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel Worksheets|*.xls;*.xlsx;*.csv"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    SetControlEnabled(false);
                    await Task.Run(() => excelModel = new ExcelModel(openFileDialog.FileName));
                    txtExcelName.Content = openFileDialog.SafeFileName;
                    txtItems.Clear();
                    SetControlEnabled(true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private async void btnExtractExcelFile_Click(object sender, RoutedEventArgs e)
        {
            if (!excelModel.IsValid())
            {
                ErrorMessageBox.Show("Please select a valid Excel file!");
                return;
            }

            try
            {
                SetControlEnabled(false);

                await ProcessCellValuesAsync(excelModel.Worksheet, excelModel.Range);

                excelModel.Workbook.Close(true, null, null);
                excelModel.Application.Quit();

                Marshal.ReleaseComObject(excelModel.Worksheet);
                Marshal.ReleaseComObject(excelModel.Workbook);
                Marshal.ReleaseComObject(excelModel.Application);

                SetControlEnabled(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            foreach (string result in keysText)
            {
                txtItems.Text += result;
                txtItems.Text += Environment.NewLine;
            }

            txtStartingRowNumber.IsEnabled = false;
            txtColumnNumber.IsEnabled = false;
            btnExtractExcelFile.IsEnabled = false;
        }

        private async Task ProcessCellValuesAsync(Excel.Worksheet worksheet, Excel.Range range)
        {
            int startingRowNumber = 0, columnNumber = 0;

            try
            {
                startingRowNumber = int.Parse(txtStartingRowNumber.Text);
                columnNumber = int.Parse(txtColumnNumber.Text);
            }
            catch(Exception ex)
            {
                ErrorMessageBox.Show(ex.Message);
                return;
            }

            for (int rowCount = startingRowNumber; rowCount <= range.Rows.Count; rowCount++)
            {
                string value = await Task.Run(() => GetCellValue(worksheet, rowCount, columnNumber));
                keysText.Add(value);
            }
        }

        private string GetCellValue(Excel.Worksheet worksheet, int row, int column) => worksheet.Cells[row, column].Text.ToString();

        private void btnSearchDirectory_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                txtSearchDirectory.Text = dialog.FileName;
            }
        }

        private async void btnStartSearch_Click(object sender, RoutedEventArgs e)
        {
            if (txtItems.Text == string.Empty)
            {
                ErrorMessageBox.Show("Please select a valid Excel file.");
                return;
            }

            if (txtSearchDirectory.Text == string.Empty)
            {
                ErrorMessageBox.Show("Please enter a valid search directory path.");
                return;
            }

            SetControlEnabled(false);

            await FindItemsAsync();

            SetControlEnabled(true);

            txtSearchResults.Clear();

            foreach (string result in resultText)
            {
                txtSearchResults.Text += result;
                txtSearchResults.Text += Environment.NewLine;
            }

            MessageBox.Show("Search has been successful!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private async Task FindItemsAsync()
        {
            string dirScanner = txtSearchDirectory.Text;

            if (keysText.Count < 1)
            {
                string str = txtItems.Text;
                keysText = str.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries).ToList();
            }

            string[] filesToSearch = await GetFilesAsync(dirScanner, txtFileExtensions.Text);

            string prefix = txtPrefix.Text, suffix = txtSuffix.Text;

            foreach (string key in keysText)
            {
                if (string.IsNullOrWhiteSpace(key))
                {
                    break;
                }
                
                for(int count = 0; count < filesToSearch.Length; count++)
                {
                    string firstOccurrence = await FindItemOccurence(filesToSearch[count], key, prefix, suffix);

                    if (firstOccurrence != null)
                    {
                        resultText.Add("Used");
                        break;
                    }

                    if (count == filesToSearch.Length - 1)
                    {
                        resultText.Add("Not Used");
                    }
                }
            }
        }

        private async Task<string> FindItemOccurence(string fileExtension, string key, string prefix, string suffix)
        {
            string[] lines = File.ReadAllLines(fileExtension);
            return await Task.Run(() => lines.FirstOrDefault(l => l.Contains(prefix + key + suffix)));
        }

        // Helper Methods

        public static async Task<string[]> GetFilesAsync(string path, string searchPattern)
        {
            string[] searchPatterns = searchPattern.Split('|');
            List<string> files = new List<string>();
            
            foreach (string sp in searchPatterns)
            {
                await Task.Run(() => files.AddRange(Directory.GetFiles(path, sp, SearchOption.AllDirectories)));
            }

            files.Sort();
            return files.ToArray();
        }

        public void SetControlEnabled(bool value)
        {
            foreach (Control control in controls)
            {
                control.IsEnabled = value;
            }
        }
    }
}
