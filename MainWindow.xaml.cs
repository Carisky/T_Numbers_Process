using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Storage.Pickers;
using WinRT.Interop;
using ClosedXML.Excel;

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
            }
        }

        private void OnStartClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_sourceFilePath) || string.IsNullOrEmpty(_fileToCheckPath))
            {
                return;
            }

            var sourceValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var sourceWb = new XLWorkbook(_sourceFilePath))
            {
                var ws = sourceWb.Worksheets.First();
                foreach (var cell in ws.Column("A").CellsUsed(c => c.Address.RowNumber >= 2))
                {
                    var text = cell.GetString();
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        sourceValues.Add(text.Trim());
                    }
                }
            }

            using (var checkWb = new XLWorkbook(_fileToCheckPath))
            {
                var ws = checkWb.Worksheets.First();
                foreach (var cell in ws.Column("C").CellsUsed(c => c.Address.RowNumber >= 2))
                {
                    var text = cell.GetString();
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    if (sourceValues.Contains(text.Trim()))
                    {
                        cell.Style.Font.FontColor = XLColor.Green;
                    }
                    else
                    {
                        cell.Style.Font.FontColor = XLColor.Red;
                    }
                }

                checkWb.Save();
            }
        }
    }
}
