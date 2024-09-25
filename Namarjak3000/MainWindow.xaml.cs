using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using OfficeOpenXml;
using Ookii.Dialogs.Wpf;

namespace Namarjak3000;

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
    private bool isEnglish = true;

    private void LanguageToggleButton_Click(object sender, RoutedEventArgs e)
    {
        isEnglish = !isEnglish;
        UpdateLanguage();
    }

    private void UpdateLanguage()
    {
        if (isEnglish)
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

    private void UpdateProgress(int current, int total) => Dispatcher.Invoke(() =>
    {
        ProgressBar.Value = (double)current / total * 100;
        ProgressLabel.Text = $"Progress: {current}/{total} rows processed.";
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

    private void ShowError(string message)
    {
        MessageBox.Show(message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    // Allow window dragging
    private void Window_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton == MouseButton.Left)
            this.DragMove();
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
        var dialog = new VistaFolderBrowserDialog();

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        _outputFolder = dialog.SelectedPath;
        Log($"Output folder set to: {_outputFolder}");
    }

// Generate Word documents with a streaming approach
    private async void GenerateDocuments_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_excelFilePath) || string.IsNullOrEmpty(_wordTemplatePath) ||
            string.IsNullOrEmpty(_outputFolder))
        {
            MessageBox.Show("Please select Excel files, Word templates, and an output folder.", "Error",
                MessageBoxButton.OK, MessageBoxImage.Error);
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
                int totalRows = 0;
                totalRows += CountRowsInExcel(_excelFilePath);

                int processedRows = 0;
                var progress = new Progress<int>(value => UpdateProgress(value, totalRows));

                string outputFilePath = Path.Combine(_outputFolder,
                    $"{Path.GetFileNameWithoutExtension(_wordTemplatePath)}_Result.docx");
                Log($"Processing template: {_wordTemplatePath} with Excel file: {_excelFilePath}");

                processedRows +=
                    ProcessTemplateInChunks(_wordTemplatePath, outputFilePath, _excelFilePath, 10, progress);
                Log($"Generated document: {outputFilePath}");
            });

            MessageBox.Show("Documents generated successfully!", "Success", MessageBoxButton.OK,
                MessageBoxImage.Information);
            Log("Document generation completed successfully.");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK,
                MessageBoxImage.Error);
            Log($"Error during document generation: {ex}");
        }
        finally
        {
            GenerateDocumentsButton.IsEnabled = true;
        }
    }

// Count the number of rows in the Excel file
    private int CountRowsInExcel(string excelFilePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            return worksheet.Rows.Count() - 1; // Subtract header row
        }
    }

    static int ProcessTemplateInChunks(string templatePath, string outputPath, string excelFilePath, int chunkSize,
        IProgress<int> progress)
    {
        int processedRows = 0;

        using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputPath,
                   DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body;

            var data = ReadExcelDataInChunks(excelFilePath, chunkSize).ToList();
            foreach (var chunk in data)
            {
                foreach (var rowData in chunk)
                {
                    // Clone the entire template
                    var clonedTemplate = CloneAndReplacePlaceholders(GetTemplateBody(templatePath), rowData);
                    body.Append(clonedTemplate);
                    body.Append(new Paragraph(new Run(new Break()
                        { Type = BreakValues.Page }))); // Page break after each clone

                    processedRows++;
                }

                // Update progress after processing each chunk
                progress.Report(processedRows);
            }

            mainPart.Document.Save();
        }

        return processedRows; // Return the number of processed rows
    }

// Method to read the template body
    static Body GetTemplateBody(string templatePath)
    {
        using (WordprocessingDocument doc = WordprocessingDocument.Open(templatePath, false))
        {
            return (Body)doc.MainDocumentPart.Document.Body.CloneNode(true);
        }
    }

// Method to read data from Excel in chunks
    static IEnumerable<List<Dictionary<string, string>>> ReadExcelDataInChunks(string excelFilePath, int chunkSize)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            var columnNames = worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column]
                .Select(cell => cell.Text.Trim())
                .ToList();

            for (int row = 2; row <= worksheet.Dimension.End.Row; row += chunkSize)
            {
                var chunk = new List<Dictionary<string, string>>();

                for (int innerRow = row;
                     innerRow < row + chunkSize && innerRow <= worksheet.Dimension.End.Row;
                     innerRow++)
                {
                    var rowData = new Dictionary<string, string>();
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        rowData[columnNames[col - 1]] = worksheet.Cells[innerRow, col].Text.Trim();
                    }

                    chunk.Add(rowData);
                }

                Log($"Read chunk of {chunk.Count} rows from Excel.");
                yield return chunk; // Return the entire chunk
            }
        }
    }

// Method to clone the template and replace placeholders
    static Body CloneAndReplacePlaceholders(Body templateBody, Dictionary<string, string> replacements)
    {
        var clonedBody = (Body)templateBody.CloneNode(true);

        foreach (var paragraph in clonedBody.Descendants<Paragraph>())
        {
            foreach (var run in paragraph.Descendants<Run>())
            {
                foreach (var text in run.Descendants<Text>())
                {
                    foreach (var key in replacements.Keys)
                    {
                        if (text.Text.Contains($"@{key}"))
                        {
                            text.Text = text.Text.Replace($"@{key}", replacements[key]);
                        }
                    }
                }
            }
        }

        return clonedBody;
    }
}