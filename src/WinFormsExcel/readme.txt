**** WinForms Excel Library (.NET Component) ****
Developer: MSDN.WhiteKnight
Web site: http://smallsoft2.blogspot.ru/

Requirements:
.NET Framework 4.0+ or .NET Core 3.1+ (must be present on target machine)
MS Excel 2003+ (must be present on target machine)
Visual Studio 2010+ (for .NET Framework) or Visual Studio 2019+ (for .NET Core)

--- Overview ---
WinForms Excel Library allows to host MS Excel interface in Windows Forms application as a user control in order to display and edit one or more tables of data. The Excel application is started in different process, but the interface is fully integrated into your application as for any usual control. Basically, it works like DataGridView, but enables user to take advantage of various data processing and visualisation capabilities of MS Excel, which might be useful in scientific, engineering or financial applications. The library interacts with Excel via Primary Interop Assemblies. Currently implemented functionality:
- Filling data from files, DataTable objects or manually on cell-by-cell basis;
- Displaying and editing one or several tables in your form in Excel GUI;
- Getting data back as DataTable objects or getting the content of specific cells;
- Saving data into file;
- Adding charts;
- Manipulating worksheets: adding, deleting, reordering, changing names, changing active sheet;
- Showing/hiding Excel formula bar and status bar;
- Direct access to Excel interoperability interfaces via Application object.
The library distribution package contains demo application project (Visual Studio 2010)

--- Usage ---
The library functionality is provided via ExtraControls.AdvancedDataGrid user control. In order to use it in WinForms application, do the following:

1. Copy WinFormsExcel.dll, Interop.Excel.dll, Interop.Microsoft.Office.Core.dll assemblies in your project, and add reference to WinFormsExcel assembly.
2. Open Windows Forms designer, right-click toolbox, click "Choose elements". Click "Browse", and choose WinFormsExcel.dll. AdvancedDataGrid will appear in "NET Components" list.
3. Check AdvancedDataGrid and click OK. AdvancedDataGrid will appear in toolbox under "Other" category.
4. Drag AdvancedDataGrid onto the form.
5. Adjust control's appearance properties such as border, background color, to distinguish it visually from form's background.
Note: control will look empty in designer, because Excel is not initialized yet.
6. Place advancedDataGrid.InitializeExcel() method call in your form's constructor or OnLoad event (this will create Excel process)
7. Fill the grid with data using one of the following methods:

public void OpenFile(string file);
public void SetCellContent(int sheet, int row, int col, object val);
public void SetSheetContent(int sheet, DataTable t);

Receive user input via one of the following methods:
public DataTable GetSheetContent(int sheet, bool FirstRowHasHeaders,int n_col=0,int n_row=0);
public object GetCellContent(int sheet, int row, int col);

Use other functionality, such as manipulating sheets, adding charts, or saving data to file.

When you no longer need the control (for example, in FormClosing event), call advancedDataGrid1.Destroy() method to clean up resources.
