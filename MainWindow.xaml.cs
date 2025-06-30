using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Storage.Pickers;
using WinRT.Interop;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace T_Numbers_Check
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        private string? _sourceFilePath;
        private string? _fileToCheckPath;
        private int _sourceSheetIndex = 0;
        private int _checkSheetIndex = 0;
        public MainWindow()
        {
            InitializeComponent();
        }

        private async void OnAddSourceFileClick(object sender, RoutedEventArgs e)
        {
            var picker = new FileOpenPicker();
            picker.FileTypeFilter.Add(".xlsx");
            picker.FileTypeFilter.Add(".xls");
            InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));
            var file = await picker.PickSingleFileAsync();
            if (file != null)
            {
                _sourceFilePath = file.Path;
                using var stream = new FileStream(_sourceFilePath, FileMode.Open, FileAccess.Read);
                var wb = WorkbookFactory.Create(stream);
                var names = Enumerable.Range(0, wb.NumberOfSheets).Select(i => wb.GetSheetName(i)).ToList();
                SourceSheetComboBox.ItemsSource = names;
                SourceSheetComboBox.SelectedIndex = 0;
                SourceSheetComboBox.Visibility = Visibility.Visible;
            }
        }

        private async void OnAddFileToCheckClick(object sender, RoutedEventArgs e)
        {
            var picker = new FileOpenPicker();
            picker.FileTypeFilter.Add(".xlsx");
            picker.FileTypeFilter.Add(".xls");
            InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));
            var file = await picker.PickSingleFileAsync();
            if (file != null)
            {
                _fileToCheckPath = file.Path;
                using var stream = new FileStream(_fileToCheckPath, FileMode.Open, FileAccess.Read);
                var wb = WorkbookFactory.Create(stream);
                var names = Enumerable.Range(0, wb.NumberOfSheets).Select(i => wb.GetSheetName(i)).ToList();
                CheckSheetComboBox.ItemsSource = names;
                CheckSheetComboBox.SelectedIndex = 0;
                CheckSheetComboBox.Visibility = Visibility.Visible;
            }
        }

        private async void OnStartClick(object sender, RoutedEventArgs e)
        {
            _sourceSheetIndex = SourceSheetComboBox.SelectedIndex;
            _checkSheetIndex = CheckSheetComboBox.SelectedIndex;

            if (string.IsNullOrEmpty(_sourceFilePath) || string.IsNullOrEmpty(_fileToCheckPath))
                return;

            StatusText.Text = "Processing...";
            StartButton.IsEnabled = false;

            await Task.Run(() =>
            {
                var sourceValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                using (var sourceStream = new FileStream(_sourceFilePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook sourceWb = WorkbookFactory.Create(sourceStream);
                    var sheet = sourceWb.GetSheetAt(_sourceSheetIndex);
                    for (int r = 1; r <= sheet.LastRowNum; r++)
                    {
                        var row = sheet.GetRow(r);
                        if (row == null) continue;

                        var cell = row.GetCell(0);
                        if (cell == null) continue;

                        var text = cell.ToString();
                        if (!string.IsNullOrWhiteSpace(text))
                            sourceValues.Add(text.Trim());
                    }
                }

                IWorkbook checkWb;
                using (var checkRead = new FileStream(_fileToCheckPath, FileMode.Open, FileAccess.Read))
                {
                    checkWb = WorkbookFactory.Create(checkRead);
                }

                var checkSheet = checkWb.GetSheetAt(_checkSheetIndex);

                var greenFont = checkWb.CreateFont();
                greenFont.FontName = "Calibri";
                greenFont.FontHeightInPoints = 11;
                greenFont.Color = IndexedColors.Green.Index;
                greenFont.Boldweight = (short)FontBoldWeight.Bold;

                var greenStyle = checkWb.CreateCellStyle();
                greenStyle.SetFont(greenFont);
                greenStyle.FillForegroundColor = IndexedColors.LightGreen.Index;
                greenStyle.FillPattern = FillPattern.SolidForeground;
                greenStyle.FillBackgroundColor = IndexedColors.LightGreen.Index;

                var redFont = checkWb.CreateFont();
                redFont.FontName = "Calibri";
                redFont.FontHeightInPoints = 11;
                redFont.Color = IndexedColors.Red.Index;
                redFont.Boldweight = (short)FontBoldWeight.Bold;

                var redStyle = checkWb.CreateCellStyle();
                redStyle.SetFont(redFont);
                redStyle.FillForegroundColor = IndexedColors.Rose.Index;
                redStyle.FillPattern = FillPattern.SolidForeground;
                redStyle.FillBackgroundColor = IndexedColors.Rose.Index;

                for (int r = 1; r <= checkSheet.LastRowNum; r++)
                {
                    var row = checkSheet.GetRow(r);
                    if (row == null) continue;

                    var cell = row.GetCell(2);
                    if (cell == null) continue;

                    var text = cell.ToString();
                    if (string.IsNullOrWhiteSpace(text)) continue;

                    if (sourceValues.Contains(text.Trim()))
                        cell.CellStyle = greenStyle;
                    else
                        cell.CellStyle = redStyle;
                }

                using (var writeStream = new FileStream(_fileToCheckPath, FileMode.Create, FileAccess.Write))
                {
                    checkWb.Write(writeStream);
                }
            });

            StatusText.Text = "Done.";
            StartButton.IsEnabled = true;
        }

    }
}
