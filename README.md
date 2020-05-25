# WinForms Excel Library (.NET Component) 

[![Nuget](https://img.shields.io/nuget/v/WinFormsExcel)](https://www.nuget.org/packages/WinFormsExcel/)

**License:** [BSD 3-Clause](LICENSE)

## Requirements

- Visual Studio 2010+
- .NET Framework 4.0+ (must be present on target machine)
- MS Excel 2003+ (must be present on target machine)

## Overview

WinForms Excel Library allows to host MS Excel interface in Windows Forms application as a user control in order to display and edit one or more tables of data. The Excel application is started in different process, but the interface is fully integrated into your application as for any usual control. Basically, it works like DataGridView, but enables user to take advantage of various data processing and visualization capabilities of MS Excel, which might be useful in scientific, engineering or financial applications. The library interacts with Excel via Primary Interop Assemblies. Currently implemented functionality:

- Filling data from files, DataTable objects or manually on cell-by-cell basis;
- Displaying and editing one or several tables in your form in Excel GUI;
- Getting data back as DataTable objects or getting the content of specific cells;
- Saving data into file;
- Adding charts;
- Manipulating worksheets: adding, deleting, reordering, changing names, changing active sheet;
- Showing/hiding Excel formula bar and status bar;
- Direct access to Excel interoperability interfaces via Application object.

The library distribution package contains demo application project (Visual Studio 2010)

## Errors

*System.Runtime.InteropServices.COMException (HRESULT 0x800A03EC) when trying to set cell content.* This error happens when you try to programmatically set the content of the cell that user is currently editing (i.e., when user double clicked the cell and Excel entered edit mode on it, which displays input caret inside cell and enables user to type text). The only current workaround is to catch exception an ask user to exit from edit more by single-clicking on some other cell.
