ISpreadsheet.net
================

Excel handling library that abstracts both xls and xlsx formats. It uses the excelent EPPlus and NPOI projects
to handle the specific file formats. It can work with both files and streams (useful for web uploading).

Building the project
--------------------

You will need Visual Studio 2010 SP1 and the latest NuGet Package Manager version installed.
Once you have this, it will download the dependencies in the first build.
All the output goes to the 'build' folder and if you want to generate documentation, just switch
your configuration to 'Release'.

Basic Usage
-----------

To use ISpreadsheet in your project you will have to reference the ISpreadsheet.dll assembly and
add references to both EPPlus and NPOI, either manually or using NuGet to get the packages online.

In order to open an existing workbook, you need to use one of the static methods from SpreadsheetFactory,
depending if you want to open a file or a stream, like this;

```C#
// for files
var book = SpreadsheetFactory.GetWorkbook("book1.xlsx");

//for streams
// ... inputStream is a stream of some kind (MemoryStream, HttpPostedFile, etc.)
var book = SpreadsheetFactory.GetWorkbook(inputStream, "xls");
```
The stream overload also takes the file extension in order to resolve which concrete handler it will use.

Open worksheets and read Values
-------------------------------

To read values from an existing worksheet you can use one of this methods:

* GetSheet(string name)
* GetSheet(int num)

Or iterate over the Sheets array to process every sheet.

Once you have obtained an IWorksheet from the book, you can use the GetCell() or GetString() methods to read
values from it, using either column/row numerical indexes or simple Excel-like addresses such as "A1" or "C25".

You could also use any of the provided extension methods in WorksheetExtensions to try to read different value types 
from a cell such as int?, float?, double?, decimal? and DateTime?.