# NamarjakProX - Word Template Copier with Excel Data Integration

### Overview

**NamarjakProX** is a console application that copies a Word template (.docx) file multiple times, replacing placeholders in the template with data from an Excel file (.xlsx). It uses the **Open XML SDK** for working with Word documents and Excel files.

The application reads an Excel file where the first row contains column headers (which are used as placeholders in the Word template) and the subsequent rows contain data. For each row in the Excel file, a new copy of the Word template is created with the placeholders replaced by the corresponding data from the Excel file.

### How It Works

1. **Input Files**: 
   - You provide the path to an Excel file (`.xlsx`) and a Word template (`.docx`).
   - The Word template contains placeholders formatted as `@PlaceholderName` that correspond to column headers in the Excel file.
   
2. **Placeholder Replacement**:
   - The application reads the Excel file, identifies the headers in the first row, and replaces the placeholders in the Word document with the corresponding data from each row of the Excel file.

3. **Output**:
   - For each row of data in the Excel file, the app creates a new copy of the Word template with the placeholders replaced.
   - All copies are saved in a folder named after the original Word template.
   
4. **Zipping**:
   - After generating the copies, the application creates a zip archive containing all the Word documents.

### Example

- **Excel File Structure**:
  - The Excel file should have a structure where the first row contains headers (e.g., `Name`, `Address`, `Date`), and subsequent rows contain the data.
  
  | Name     | Address         | Date       |
  | -------- | --------------- | ---------- |
  | John Doe | 123 Main St.     | 01/01/2024 |
  | Jane Roe | 456 Elm St.      | 02/01/2024 |
  
- **Word Template**:
  - In your Word template, use placeholders that match the column headers in the Excel file, formatted as `@Name`, `@Address`, and `@Date`.
  
-  **Template Example**:
  Dear @Name,
  Your address is: @Address The date is: @Date

- **Output**:
- The app will generate multiple Word documents with the placeholders replaced by actual data from the Excel file.

### Installation

1. **Build the Application**:
 - This is a .NET Core console application. You can build it by using the following command in your project directory:
   ```bash
   dotnet build
   ```

2. **Running the Application**:
 - After building, you can run the executable from the command line. The application will prompt you for the path to the Excel file and the Word template.

### Requirements

- .NET Core SDK installed on your machine.
- A `.docx` Word template with placeholders in the format `@PlaceholderName`.
- A `.xlsx` Excel file where the first row contains the headers (corresponding to placeholders) and subsequent rows contain data.

### Instructions

1. **Prepare Your Word Template**:
 - In your Word template, use placeholders such as `@Name`, `@Address`, etc., corresponding to the headers in your Excel file.

2. **Prepare Your Excel File**:
 - The first row should contain headers that match your placeholders in the Word document (without the `@` symbol).
 - Each subsequent row should contain data that will replace the placeholders in the Word template.

3. **Run the Application**:
 - After building the application, execute the `.exe` file or run it from the command line using `dotnet run`.
 - You will be prompted to input the paths for the Excel file and Word template.

4. **Output**:
 - The application will create a folder with copies of the Word template, where placeholders have been replaced by data from the Excel file. The folder will be zipped automatically.

### Example Usage

```bash
dotnet run
```
# License
The Open XML SDK is licensed under the [MIT]("https://github.com/bzorec1/namarjak3000?tab=MIT-1-ov-file") license.

# Notice

This license content is provided by the Open XML SDK project. For more information about licenses.nuget.org, see [our documentation]("https://learn.microsoft.com/en-us/nuget/nuget-org/licenses.nuget.org").
