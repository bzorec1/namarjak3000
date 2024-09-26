using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32;

namespace Namarjak3000;

// ReSharper disable once RedundantExtendsListEntry
// ReSharper disable once UnusedMember.Global
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        UpdateLanguage();
    }

    private string? _excelFilePath;
    private string? _wordTemplatePath;
    private string? _outputFolder;
    private int _processedRows;
    private int _totalRows;
    private bool _isEnglish = true;

    private void LanguageToggleButton_Click(object sender, RoutedEventArgs e)
    {
        _isEnglish = !_isEnglish;
        UpdateLanguage();
    }

    private void UpdateLanguage()
    {
        if (_isEnglish)
        {
            LanguageToggleButton.Content = "EN";
            WelcomeText.Text = "Welcome to Excel to Word Generator!";
            ProgramPurpose.Text =
                "This tool helps you generate Word documents by combining Excel data with Word templates.";
            Step1Text.Text = "1. Add your Excel file (Only one at a time)";
            Step2Text.Text = "2. Add a Word template with @placeholders (no spaces in placeholders)";
            Step3Text.Text = "3. Select an output folder for the generated documents";
            HeaderInstructionText.Text = "Ensure the Excel headers match the Word @placeholders exactly (no spaces).";
        }
        else
        {
            LanguageToggleButton.Content = "SI";
            WelcomeText.Text = "Dobrodošli v generatorju Excel v Word!";
            ProgramPurpose.Text =
                "To orodje vam pomaga ustvariti Word dokumente z združevanjem podatkov iz Excela in Word predlog.";
            Step1Text.Text = "1. Dodajte svojo Excel datoteko (le eno naenkrat)";
            Step2Text.Text = "2. Dodajte Word predlogo z @mesto-za-vstavljanje (brez presledkov v imenu)";
            Step3Text.Text = "3. Izberite izhodno mapo za ustvarjene dokumente";
            HeaderInstructionText.Text =
                "Poskrbite, da se glave Excel ujemajo z Word @mesto-za-vstavljanje natančno (brez presledkov).";
        }
    }

    private void MinimizeButton_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;

    private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();

    private static void Log(string message) => Debug.WriteLine($"[{DateTime.Now}] {message}");

    private void UpdateProgress() => Dispatcher.Invoke(() =>
    {
        ProgressBar.Value = (double)_processedRows / _totalRows;
        ProgressLabel.Text = $"Progress: {_processedRows}/{_totalRows} rows processed.";
    });

    private void AddExcelFiles_Click(object sender, RoutedEventArgs e)
    {
        var openFileDialog = new OpenFileDialog
        {
            Filter = "Excel files (*.xlsx)|*.xlsx"
        };

        if (openFileDialog.ShowDialog() != true)
        {
            return;
        }

        var file = openFileDialog.FileNames.FirstOrDefault();
        _excelFilePath = file ?? throw new InvalidOperationException();

        if (!_excelFilePath.Equals(file))
        {
            Log($"Added Excel file: {file}");
        }
    }

    private void ShowError(string message) =>
        MessageBox.Show(message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);

    private void Window_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton == MouseButton.Left)
        {
            DragMove();
        }
    }

    private void AddWordTemplates_Click(object sender, RoutedEventArgs e)
    {
        var openFileDialog = new OpenFileDialog
        {
            Filter = "Word files (*.docx)|*.docx"
        };

        if (openFileDialog.ShowDialog() != true)
        {
            return;
        }

        var file = openFileDialog.FileNames.FirstOrDefault();
        _wordTemplatePath = file ?? throw new InvalidOperationException();

        if (!_wordTemplatePath.Equals(file))
        {
            Log($"Added Word template: {file}");
        }
    }

    private void BrowseOutputFolder_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog();

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        _outputFolder = dialog.DefaultDirectory;
        Log($"Output folder set to: {_outputFolder}");
    }

    private async void GenerateDocuments_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_excelFilePath) || string.IsNullOrEmpty(_wordTemplatePath) ||
            string.IsNullOrEmpty(_outputFolder))
        {
            ShowError("Please select Excel files, Word templates, and an output folder.");
            return;
        }

        GenerateDocumentsButton.IsEnabled = false;
        ProgressBar.Value = 0;
        ProgressLabel.Text = "Progress: 0/0 rows processed.";
        Log("Starting document generation...");

        try
        {
            await Task.Run(() =>
            {
                var excelData = ReadExcelFile(_excelFilePath);
                _totalRows = excelData.Sum(i => i.Value.Count) / excelData.Count;

                UpdateProgress();

                var outputFilePath = Path.Combine(_outputFolder, $"{Path.GetFileNameWithoutExtension(_wordTemplatePath)}_Result.docx");

                for (int i = 0; i < _totalRows; i++)
                {
                    _processedRows++;
                    UpdateProgress();
                }
            });

            MessageBox.Show("Documents generated successfully!", "Success", MessageBoxButton.OK,
                MessageBoxImage.Information);
            Log("Document generation completed successfully.");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            Log($"Error during document generation: {ex}");
        }
        finally
        {
            GenerateDocumentsButton.IsEnabled = true;
        }
    }

    public static Dictionary<string, List<string>> ReadExcelFile(string excelFilePath)
    {
        var headerDictionary = new Dictionary<string, List<string>>();

        using var document = SpreadsheetDocument.Open(excelFilePath, false);

        var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException();
        var sheet = workbookPart.Workbook.Sheets?.GetFirstChild<Sheet>() ?? throw new InvalidOperationException();
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>() ?? throw new InvalidOperationException();

        var headerRow = sheetData.Elements<Row>().First();
        foreach (var cell in headerRow.Elements<Cell>())
        {
            var header = GetCellValue(document, cell);
            headerDictionary[header] = [];
        }

        foreach (var row in sheetData.Elements<Row>().Skip(1))
        {
            var columnIndex = 0;
            foreach (var cell in row.Elements<Cell>())
            {
                var header = GetCellValue(document, headerRow.Elements<Cell>().ElementAt(columnIndex));
                var value = GetCellValue(document, cell);

                if (headerDictionary.ContainsKey(header))
                {
                    headerDictionary[header].Add(value);
                }

                columnIndex++;
            }
        }

        return headerDictionary;
    }

    private static string GetCellValue(SpreadsheetDocument document, Cell cell)
    {
        if (cell == null || string.IsNullOrEmpty(cell.InnerText))
        {
            return string.Empty;
        }

        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            return document.WorkbookPart?.SharedStringTablePart?.SharedStringTable.Elements<SharedStringItem>()
                .ElementAt(int.Parse(cell.InnerText)).InnerText ?? throw new InvalidOperationException();
        }

        return cell.InnerText;
    }
}