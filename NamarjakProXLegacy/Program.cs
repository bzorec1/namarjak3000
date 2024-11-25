using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WP = DocumentFormat.OpenXml.Wordprocessing;

namespace NamarjakProX;

public class Program
{
    private static int _totalRows;
    private static bool _debug;

    public static async Task Main(string[] args)
    {
        Console.WriteLine("Welcome to NamarjakProXLegacy - Template Document Copier");
        Console.WriteLine("This application copies a template Word document and replaces placeholders based on an Excel file.");
        Console.WriteLine();

        Console.WriteLine("Enable debug mode? (y/n):");
        string debugInput = Console.ReadLine()?.ToLower() ?? throw new InvalidOperationException();
        _debug = debugInput == "y" || debugInput == "yes";

        PrintInstructions();

        while (true)
        {
            Console.WriteLine("\nPress 'q' to quit or any other key to continue.");
            string quitInput = Console.ReadLine()?.ToLower() ?? string.Empty;
            if (quitInput == "q")
            {
                Console.WriteLine("Exiting program. Goodbye!");
                break;
            }

            string excelFilePath = string.Empty;
            string sourceFilePath = string.Empty;

            while (string.IsNullOrEmpty(excelFilePath) || string.IsNullOrEmpty(sourceFilePath))
            {
                if (string.IsNullOrEmpty(excelFilePath))
                {
                    Console.WriteLine("Enter the path to the Excel file:");
                    excelFilePath = Console.ReadLine()?.Trim('\"') ?? string.Empty;

                    if (excelFilePath.ToLower() == "q")
                    {
                        Console.WriteLine("NamarjakProXLegacy se je zmantro!");
                        return;
                    }
                }

                if (string.IsNullOrEmpty(sourceFilePath))
                {
                    Console.WriteLine("Enter the path to the source file (template .docx):");
                    sourceFilePath = Console.ReadLine()?.Trim('\"') ?? string.Empty;

                    if (sourceFilePath.ToLower() == "q")
                    {
                        Console.WriteLine("Exiting program. Goodbye!");
                        return;
                    }
                }

                try
                {
                    ValidateFilePaths(excelFilePath, sourceFilePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                    excelFilePath = string.Empty;
                    sourceFilePath = string.Empty;
                }
            }

            string baseFileName = Path.GetFileNameWithoutExtension(sourceFilePath);
            string destinationDirectory = PrepareDestinationDirectory(sourceFilePath, baseFileName);

            Console.WriteLine("Starting to copy and process files...");
            using (var progressBar = new ProgressBar())
            {
                try
                {
                    Dictionary<string, List<string>> excelData = ReadExcelFile(excelFilePath);
                    _totalRows = excelData.Max(i => i.Value.Count(x => !string.IsNullOrEmpty(x)));

                    await ProcessDocumentsAsync(baseFileName, sourceFilePath, destinationDirectory, excelData,
                        progressBar);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Critical error: {ex.Message}");
                }
            }

            Console.WriteLine("\nZipping the folder...");
            string zipFilePath = ZipFolder(destinationDirectory, sourceFilePath);
            Console.WriteLine($"Zipped the folder to: {zipFilePath}");

            Console.Clear();
        }
    }

    private static void PrintInstructions()
    {
        Console.WriteLine("Instructions:");
        Console.WriteLine("1. Prepare an Excel file with headers as placeholders (e.g., @Name, @Date, @Address).");
        Console.WriteLine("2. Prepare a Word document template with placeholders matching the Excel headers.");
        Console.WriteLine("3. Save the Excel file in .xlsx format and the Word document in .docx format.");
        Console.WriteLine("   Note their file paths for the next steps.\n");
    }

    private static void ValidateFilePaths(string excelFilePath, string sourceFilePath)
    {
        if (!string.IsNullOrEmpty(excelFilePath) && !File.Exists(excelFilePath))
        {
            throw new FileNotFoundException($"The Excel file '{excelFilePath}' does not exist.");
        }

        if (!string.IsNullOrEmpty(sourceFilePath) && !File.Exists(sourceFilePath))
        {
            throw new FileNotFoundException($"The source file '{sourceFilePath}' does not exist.");
        }
    }

    private static string PrepareDestinationDirectory(string sourceFilePath, string baseFileName)
    {
        string destinationDirectory = Path.Combine(Path.GetDirectoryName(sourceFilePath) ?? string.Empty, baseFileName);
        Directory.CreateDirectory(destinationDirectory);

        foreach (var file in Directory.GetFiles(destinationDirectory))
        {
            File.Delete(file);
        }

        return destinationDirectory;
    }

    private static async Task ProcessDocumentsAsync(string baseFileName, string sourceFilePath,
        string destinationDirectory, Dictionary<string, List<string>> excelData, ProgressBar progressBar)
    {
        string fileExtension = Path.GetExtension(sourceFilePath);
        int progressCount = 0;

        await Task.Run(() =>
        {
            for (int i = 0; i < _totalRows; i++)
            {
                try
                {
                    string newFileName = $"{baseFileName}_{i + 1}{fileExtension}";
                    string destinationFilePath = Path.Combine(destinationDirectory, newFileName);

                    File.Copy(sourceFilePath, destinationFilePath, overwrite: false);

                    using (WordprocessingDocument document = WordprocessingDocument.Open(destinationFilePath, true))
                    {
                        MainDocumentPart mainPart =
                            document.MainDocumentPart ?? throw new InvalidOperationException();
                        ReplacePlaceholders(mainPart, excelData, i);
                        document.Save();
                    }

                    progressCount++;
                    progressBar.Report((double)progressCount / _totalRows);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing row {i + 1}: {ex.Message}");
                }
            }
        });
    }

    private static void ReplacePlaceholders(MainDocumentPart mainPart, Dictionary<string, List<string>> excelData,
        int rowIndex)
    {
        if (mainPart.Document.Body == null)
        {
            throw new InvalidOperationException("Document body is empty. Document generation failed.");
        }

        foreach (var paragraph in mainPart.Document.Body.Elements<WP.Paragraph>())
        {
            foreach (var run in paragraph.Elements<WP.Run>())
            {
                foreach (var text in run.Elements<WP.Text>())
                {
                    foreach (var header in excelData.Keys)
                    {
                        if (rowIndex < excelData[header].Count)
                        {
                            text.Text = text.Text.Replace($"@{header}", ParseData(excelData[header][rowIndex]));
                        }
                        else
                        {
                            if (_debug)
                            {
                                Console.WriteLine(
                                    $"Warning: Row index {rowIndex} is out of range for header '{header}'.");
                            }
                        }
                    }
                }
            }
        }
    }

    private static string ZipFolder(string folderPath, string sourceFilePath)
    {
        string zipFilePath = Path.Combine(Path.GetDirectoryName(sourceFilePath) ?? string.Empty,
            $"{Path.GetFileNameWithoutExtension(folderPath)}.zip");

        if (File.Exists(zipFilePath))
        {
            File.Delete(zipFilePath);
        }

        ZipFile.CreateFromDirectory(folderPath, zipFilePath);
        return zipFilePath;
    }

    private static Dictionary<string, List<string>> ReadExcelFile(string excelFilePath)
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
            headerDictionary[header] = new List<string>();
        }

        foreach (var row in sheetData.Elements<Row>().Skip(1))
        {
            try
            {
                for (int columnIndex = 0; columnIndex < headerRow.Elements<Cell>().Count(); columnIndex++)
                {
                    try
                    {
                        var header = GetCellValue(document, headerRow.Elements<Cell>().ElementAt(columnIndex));
                        var cell = row.Elements<Cell>().ElementAtOrDefault(columnIndex); // Safely get cell
                        var value = cell != null ? GetCellValue(document, cell) : string.Empty;

                        if (_debug)
                        {
                            Console.WriteLine($"Row {row.RowIndex}, Column {header}: {value}");
                        }

                        if (headerDictionary.ContainsKey(header))
                        {
                            headerDictionary[header].Add(value);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(
                            $"Error processing cell at Row {row.RowIndex}, Column {columnIndex}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error at Row {row.RowIndex}: {ex.Message}");
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
                .ElementAt(int.Parse(cell.InnerText)).InnerText ?? string.Empty;
        }

        return cell.InnerText;
    }

    private static string ParseData(string value)
    {
        if (!decimal.TryParse(value, out var numericValue))
        {
            return value;
        }

        return numericValue % 1 == 0 ? ((int)numericValue).ToString() : value;
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