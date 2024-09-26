using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using A = DocumentFormat.OpenXml.Drawing;

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
    private string? _outputFilePath;
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
        ProgressBar.Value = (double)_processedRows / _totalRows * 100;
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

        _outputFolder = dialog.FolderName;
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

        _outputFilePath = Path.Combine(_outputFolder,
            $"{Path.GetFileNameWithoutExtension(_wordTemplatePath)}_Result.docx");

        GenerateDocumentsButton.IsEnabled = false;
        _processedRows = 0;
        ProgressBar.Value = 0;
        ProgressLabel.Text = "Progress: 0/0 rows processed.";
        Log("Starting document generation...");

        try
        {
            await Task.Run(() =>
            {
                Dictionary<string, List<string>> excelData = ReadExcelFile(_excelFilePath);
                _totalRows = excelData.Max(i => i.Value.Count(x => !string.IsNullOrEmpty(x)));

                UpdateProgress();

                using (WordprocessingDocument template = WordprocessingDocument.Open(_wordTemplatePath, false))
                {
                    if (template.MainDocumentPart is null)
                    {
                        throw new Exception("Template document is missing the main part. Document generation failed.");
                    }
                
                    using (WordprocessingDocument document =
                           WordprocessingDocument.Create(_outputFilePath, WordprocessingDocumentType.Document))
                    {
                        MainDocumentPart mainPart = document.AddMainDocumentPart();
                        mainPart.Document = new Document(new Body());
                
                        for (int i = 0; i < _totalRows; i++)
                        {
                            if (mainPart.Document.Body is null)
                            {
                                throw new Exception(
                                    "Newly created document body is empty. Document generation failed. Please try again.");
                            }
                
                            List<Paragraph> currentBody = CopyMainDocumentPartContent(template, document);
                            
                            foreach (var paragraph in currentBody)
                            {
                                ReplacePlaceholdersInParagraph(paragraph, excelData, i);
                            }
                
                            if (i < _totalRows - 1)
                            {
                                AddPageBreak(mainPart);
                            }
                
                            _processedRows++;
                            UpdateProgress();
                        }
                
                        Debug.Assert(document.MainDocumentPart != null, "document.MainDocumentPart != null");
                        document.MainDocumentPart.Document.Save();
                    }
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

    private static List<Paragraph> CopyMainDocumentPartContent(WordprocessingDocument sourceDocument,
        WordprocessingDocument destinationDocument)
    {
        MainDocumentPart sourceMainPart = sourceDocument.MainDocumentPart ?? throw new InvalidOperationException();
        MainDocumentPart destinationMainPart =
            destinationDocument.MainDocumentPart ?? throw new InvalidOperationException();

        List<Paragraph> clonedParagraphs = new List<Paragraph>();
        foreach (var element in sourceMainPart.Document.Body?.Elements()!)
        {
            var clonedElement = element.CloneNode(true);
            destinationMainPart.Document.Body?.Append(clonedElement);

            if (clonedElement is Paragraph paragraph)
            {
                clonedParagraphs.Add(paragraph);
            }
        }

        return clonedParagraphs;
    }

    private static void ReplacePlaceholdersInParagraph(Paragraph paragraph, Dictionary<string, List<string>> excelData,
        int rowIndex)
    {
        foreach (var run in paragraph.Elements<Run>())
        {
            foreach (var text in run.Elements<Text>())
            {
                foreach (var header in excelData.Keys)
                {
                    text.Text = text.Text.Replace($"@{header}", ParseData(excelData[header][rowIndex]));
                }
            }
        }
    }

    private static string ParseData(string value)
    {
        if (!decimal.TryParse(value, out decimal numericValue))
        {
            return value;
        }

        return numericValue % 1 == 0 ? ((int)numericValue).ToString() : value;
    }


    private static void AddPageBreak(MainDocumentPart mainPart)
    {
        Paragraph pageBreakParagraph = new Paragraph(new Run(new Break
        {
            Type = BreakValues.Page
        }));

        mainPart.Document.Body?.AppendChild(pageBreakParagraph);
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