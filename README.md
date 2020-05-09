# WinForms Excel Library (.NET Component) 

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
