using System.Collections.Concurrent;
using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WP = DocumentFormat.OpenXml.Wordprocessing;

namespace NamarjakProX;

public class Program
{
    private static int _totalRows;

    public static async Task Main(string[] args)
    {
        Console.WriteLine("Welcome to NamarjakProX - Template Document Copier");
        Console.WriteLine(
            "This application copies a template Word document and replaces placeholders based on an Excel file.");
        Console.WriteLine();

        // Instructions for preparing Excel and Word documents
        Console.WriteLine("Instructions:");
        Console.WriteLine("1. Prepare an Excel file with the following requirements:");
        Console.WriteLine(
            "   - The first row should contain unique headers as placeholders (e.g., @Name, @Date, @Address).");
        Console.WriteLine("   - Each subsequent row should contain data corresponding to these headers.");
        Console.WriteLine("   Example:");
        Console.WriteLine("   |  Name   |  Date       |  Address      |");
        Console.WriteLine("   |---------|-------------|---------------|");
        Console.WriteLine("   | John Doe| 2024-09-26  | 123 Elm St    |");
        Console.WriteLine();
        Console.WriteLine("2. Prepare a Word document template with placeholders.");
        Console.WriteLine("   - Use the same format as in the Excel headers, prefixed by '@'.");
        Console.WriteLine("   Example content in Word:");
        Console.WriteLine("   Dear @Name,");
        Console.WriteLine("   We are pleased to inform you that your appointment is scheduled for @Date.");
        Console.WriteLine("   Please visit us at @Address.");
        Console.WriteLine();
        Console.WriteLine("3. Save the Excel file in .xlsx format and the Word document in .docx format.");
        Console.WriteLine("   Note their file paths for the next steps.\n");

        while (true)
        {
            Console.WriteLine("\nCTRL + C to exit.");
            Console.WriteLine();

            string excelFilePath = string.Empty;
            string sourceFilePath = string.Empty;

            while (string.IsNullOrEmpty(excelFilePath) || string.IsNullOrEmpty(sourceFilePath))
            {
                if (string.IsNullOrEmpty(excelFilePath))
                {
                    Console.WriteLine("Enter the path to the Excel file:");
                    excelFilePath = Console.ReadLine()?.Trim('\"') ?? string.Empty;
                }

                if (string.IsNullOrEmpty(sourceFilePath))
                {
                    Console.WriteLine("Enter the path to the source file (template .docx):");
                    sourceFilePath = Console.ReadLine()?.Trim('\"') ?? string.Empty;
                }

                try
                {
                    if (!string.IsNullOrEmpty(excelFilePath) && !File.Exists(excelFilePath))
                    {
                        Console.WriteLine($"Error: The Excel file '{excelFilePath}' does not exist.");
                        excelFilePath = string.Empty;
                    }

                    if (!string.IsNullOrEmpty(sourceFilePath) && !File.Exists(sourceFilePath))
                    {
                        Console.WriteLine($"Error: The source file '{sourceFilePath}' does not exist.");
                        sourceFilePath = string.Empty;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                    excelFilePath = string.Empty;
                    sourceFilePath = string.Empty;
                }
            }

            string baseFileName = Path.GetFileNameWithoutExtension(sourceFilePath);

            string destinationDirectory =
                Path.Combine(Path.GetDirectoryName(sourceFilePath) ?? string.Empty, baseFileName);
            Directory.CreateDirectory(destinationDirectory);

            foreach (var file in Directory.GetFiles(destinationDirectory))
            {
                File.Delete(file);
            }

            string fileExtension = Path.GetExtension(sourceFilePath);
            Console.WriteLine("Starting to copy and process files...");
            using (var progressBar = new ProgressBar())
            {
                Dictionary<string, List<string>> excelData = ReadExcelFile(excelFilePath);
                _totalRows = excelData.Max(i => i.Value.Count(x => !string.IsNullOrEmpty(x)));

                // Create a partitioner to divide the work
                var rangePartitioner = Partitioner.Create(0, _totalRows + 1);

                // Use a counter to track progress
                int progressCount = 0;

                // Run the document creation in parallel with 4 threads
                await Task.Run(() =>
                {
                    Parallel.ForEach(rangePartitioner, new ParallelOptions { MaxDegreeOfParallelism = 4 }, range =>
                    {
                        for (int i = range.Item1; i < range.Item2; i++)
                        {
                            string newFileName = $"{baseFileName}_{i + 1}{fileExtension}";
                            string destinationFilePath = Path.Combine(destinationDirectory, newFileName);

                            File.Copy(sourceFilePath, destinationFilePath, overwrite: false);

                            // Process each document
                            using (WordprocessingDocument document =
                                   WordprocessingDocument.Open(destinationFilePath, true))
                            {
                                MainDocumentPart mainPart =
                                    document.MainDocumentPart ?? throw new InvalidOperationException();

                                if (mainPart.Document.Body is null)
                                {
                                    throw new Exception(
                                        "Newly created document body is empty. Document generation failed. Please try again.");
                                }

                                foreach (var element in mainPart.Document.Body?.Elements()!)
                                {
                                    if (element is not WP.Paragraph paragraph)
                                    {
                                        continue;
                                    }

                                    foreach (var run in paragraph.Elements<WP.Run>())
                                    {
                                        foreach (var text in run.Elements<WP.Text>())
                                        {
                                            foreach (var header in excelData.Keys)
                                            {
                                                text.Text = text.Text.Replace(
                                                    $"@{header}",
                                                    ParseData(excelData[header][i]));
                                            }
                                        }
                                    }
                                }

                                document.Save();
                            }

                            // Update the progress counter in a thread-safe manner
                            Interlocked.Increment(ref progressCount);
                            progressBar.Report((double)progressCount / _totalRows);
                        }
                    });
                });
            }

            Console.WriteLine("\nZipping the folder...");

            string zipFilePath =
                Path.Combine(Path.GetDirectoryName(sourceFilePath) ?? string.Empty, $"{baseFileName}.zip");

            if (File.Exists(zipFilePath))
            {
                File.Delete(zipFilePath);
            }

            ZipFile.CreateFromDirectory(destinationDirectory, zipFilePath);
            Console.WriteLine($"Zipped the folder to: {zipFilePath}");

            Console.Clear();
        }
    }

    private static string ParseData(string value)
    {
        if (!decimal.TryParse(value, out var numericValue))
        {
            return value;
        }

        return numericValue % 1 == 0 ? ((int)numericValue).ToString() : value;
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

public class ProgressBar : IDisposable
{
    private readonly int _totalTicks;
    private int _currentTick;

    public ProgressBar(int totalTicks = 100)
    {
        _totalTicks = totalTicks;
        Console.WriteLine("Progress:");
    }

    public void Report(double value)
    {
        int progress = (int)(value * _totalTicks);
        while (_currentTick < progress)
        {
            Console.Write("█");
            _currentTick++;
        }

        Console.Write(
            $"\r[{new string('█', _currentTick)}{new string(' ', _totalTicks - _currentTick)}] {value * 100:0}%");
    }

    public void Dispose()
    {
        Console.WriteLine("\nDone!");
    }
}