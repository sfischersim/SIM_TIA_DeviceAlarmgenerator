using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using log4net;
using Microsoft.Win32;
using SIM_TIA_DeviceAlarmgenerator.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace SIM_TIA_DeviceAlarmgenerator.ViewModel
{
    internal partial class MainViewModel : ObservableObject
    {
        public ILog log = App.GetLogger();

        [ObservableProperty] private ExcelHandler excelHandler;
        [ObservableProperty] private string filePath;
        [ObservableProperty] private string startRow = "1";
        [ObservableProperty] private string endRow = "28";
        [ObservableProperty] private string startCol = "1";
        [ObservableProperty] private string endCol = "6";
        [ObservableProperty] private string sheetName= "StA-Cfg";
        [ObservableProperty] private DataView data;

        public MainViewModel()
        {
            log.Debug("Konstruktor MainViewModel aufgerufen");

            excelHandler = new ExcelHandler();
        }

        [RelayCommand]
        private void PickFile()
        {
            var dialog = new OpenFileDialog
            {
                Title = "Excel-Datei auswählen",
                Filter = "Excel-Dateien (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls",
                CheckFileExists = true
            };

            if (dialog.ShowDialog() == true)
                FilePath = dialog.FileName;

            log.Debug($"FilePath ist {dialog.FileName}");
        }

        [RelayCommand]
        private void ReadData()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(FilePath))
                {
                    MessageBox.Show("Bitte zuerst eine Excel-Datei wählen.");
                    return;
                }

                string filePath = FilePath;
                string sheetName = SheetName?.Trim();
                if (string.IsNullOrWhiteSpace(sheetName))
                {
                    MessageBox.Show("Bitte einen Sheet-Namen angeben.");
                    return;
                }

                // Eingaben parsen
                if (!int.TryParse(StartRow, out int startRow) ||
                    !int.TryParse(EndRow, out int endRow) ||
                    !int.TryParse(StartCol, out int startCol) ||
                    !int.TryParse(EndCol, out int endCol))
                {
                    MessageBox.Show("Start/Ende müssen ganze Zahlen sein.");
                    return;
                }
                if (startRow <= 0 || startCol <= 0 || endRow < startRow || endCol < startCol)
                {
                    MessageBox.Show("Bitte gültigen Bereich angeben (1-basiert, Ende ≥ Start).");
                    return;
                }

                // Lesen
                var flat = ExcelHandler.ReadRangeToArray(filePath, sheetName, startRow, endRow, startCol, endCol);

                // In DataTable für Grid umwandeln
                int rows = endRow - startRow + 1;
                int cols = endCol - startCol + 1;
                var table = ExcelHandler.FlatToDataTable(flat, rows, cols);
                Data = table.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Lesen: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
