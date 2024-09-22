using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using OfficeOpenXml;
using Ookii.Dialogs.Wpf;

namespace Namarjak3000
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private List<string> excelFilePaths = new List<string>();
        private List<string> wordTemplatePaths = new List<string>();
        private string outputFolder = "";

        private static void Log(string message)
        {
            Debug.WriteLine($"[{DateTime.Now}] {message}");
        }

        private void UpdateProgress(int current, int total)
        {
            // Use Dispatcher to update the UI from the main thread
            Dispatcher.Invoke(() =>
            {
                ProgressBar.Value = (double)current / total * 100;
                ProgressLabel.Content = $"Progress: {current}/{total} rows processed.";
            });
        }

        // Add Excel files
        private void AddExcelFiles_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Excel files (*.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                foreach (var file in openFileDialog.FileNames)
                {
                    if (!excelFilePaths.Contains(file))
                    {
                        excelFilePaths.Add(file);
                        ExcelFilesListBox.Items.Add(Path.GetFileName(file));
                        Log($"Added Excel file: {file}");
                    }
                }
            }
        }

        // Add Word templates
        private void AddWordTemplates_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Word files (*.docx)|*.docx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                foreach (var file in openFileDialog.FileNames)
                {
                    if (!wordTemplatePaths.Contains(file))
                    {
                        wordTemplatePaths.Add(file);
                        WordTemplatesListBox.Items.Add(Path.GetFileName(file));
                        Log($"Added Word template: {file}");
                    }
                }
            }
        }

        // Browse for the output folder
        private void BrowseOutputFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new VistaFolderBrowserDialog();
            if (dialog.ShowDialog() == true)
            {
                outputFolder = dialog.SelectedPath;
                OutputFolderTextBox.Text = outputFolder;
                Log($"Output folder set to: {outputFolder}");
            }
        }

        // Generate Word documents with a streaming approach
        private async void GenerateDocuments_Click(object sender, RoutedEventArgs e)
        {
            if (excelFilePaths.Count == 0 || wordTemplatePaths.Count == 0 || string.IsNullOrEmpty(outputFolder))
            {
                MessageBox.Show("Please select Excel files, Word templates, and an output folder.", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            GenerateDocumentsButton.IsEnabled = false;
            ProgressBar.Value = 0;
            ProgressLabel.Content = "Progress: 0/0 rows processed.";
            Log("Starting document generation...");

            try
            {
                await Task.Run(() =>
                {
                    int totalRows = 0;
                    totalRows += CountRowsInExcel(excelFilePaths.First());

                    int processedRows = 0;
                    var progress = new Progress<int>(value => UpdateProgress(value, totalRows));

                    foreach (var excelFile in excelFilePaths)
                    {
                        foreach (var wordTemplate in wordTemplatePaths)
                        {
                            string outputFilePath = Path.Combine(outputFolder,
                                $"{Path.GetFileNameWithoutExtension(wordTemplate)}_Result.docx");
                            Log($"Processing template: {wordTemplate} with Excel file: {excelFile}");

                            // Process the template in chunks
                            processedRows +=
                                ProcessTemplateInChunks(wordTemplate, outputFilePath, excelFile, 10, progress);
                            Log($"Generated document: {outputFilePath}");
                        }
                    }
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

        static int ProcessTemplateInChunks(string templatePath, string outputPath, string excelFilePath, int chunkSize, IProgress<int> progress)
        {
            int processedRows = 0;

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputPath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
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
                        body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page }))); // Page break after each clone

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
            Body clonedBody = (Body)templateBody.CloneNode(true);

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
}