# VBA Project Exporter

A robust VBA module for exporting all components of a VBA project (modules, classes, forms) from an Excel workbook to individual files.

## Features

- Export all VBA components from any Excel workbook
- User-friendly folder selection prompt
- Progress indicators during export
- Detailed export summary
- Handles special characters in filenames
- Option to overwrite existing files or skip them
- Well-documented code with error handling

## Installation

1. Import the `ProjectExporter.bas` module into your VBA project:
   - Open your Excel workbook
   - Press `Alt+F11` to open the VBA editor
   - Right-click on your project in the Project Explorer
   - Select `Import File...` and choose the `ProjectExporter.bas` file

2. Enable the "Microsoft Visual Basic for Applications Extensibility" library:
   - In the VBA editor, go to `Tools` > `References`
   - Check the box next to "Microsoft Visual Basic for Applications Extensibility"
   - Click OK

## Usage

### Basic Usage

1. In your Excel workbook, run the `RunExporter` procedure:
   - From the VBA editor: Select the procedure and press F5
   - From Excel: Set up a button or ribbon command to run the procedure

2. Follow the prompts to select your export directory

3. View the export summary when complete

### From VBA Code

```vba
' Export the current workbook's VBA project
Call RunExporter

' Or export from a specific workbook with custom settings
Dim success As Boolean
success = ExportVbaProject("C:\Path\To\YourWorkbook.xlsm", Workbooks("YourWorkbook.xlsm"), True)
```

## Function Reference

### `ExportVbaProject`

```vba
Public Function ExportVbaProject(ProjectPath As String, SourceWorkbook As Workbook, Optional Overwrite As Boolean = False) As Boolean
```

**Parameters:**

- `ProjectPath`: The path to the Excel workbook containing the VBA project
- `SourceWorkbook`: The workbook object containing the VBA project to export
- `Overwrite`: (Optional) Whether to overwrite existing files (Default: False)

**Returns:**

- `Boolean`: True if export was successful, False otherwise

### `RunExporter`

```vba
Public Sub RunExporter()
```

A wrapper procedure that exports the VBA project from the current workbook.

## Requirements

- Excel 2007 or later
- "Microsoft Visual Basic for Applications Extensibility" library reference
- Trust access to the VBA project object model must be enabled
  - File > Options > Trust Center > Trust Center Settings > Macro Settings > 
    Check "Trust access to the VBA project object model"

## Potential Issues

- **Trust Settings**: If trust access is not enabled, you'll receive an error
- **Locked Projects**: Password-protected VBA projects cannot be exported
- **Special Characters**: File paths with certain special characters may cause issues

## License

This code is provided under the MIT License. Feel free to modify and distribute it as needed.

## Contributing

Contributions are welcome! Please feel free to submit a pull request or open an issue if you have any suggestions or find any bugs.
