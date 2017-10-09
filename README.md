# ExcelMerge.exe

Welcome to ExcelMerge.exe. This command line utility merges all excel documents within a directory into a single excel workbook. 
This utility has a very specific use case - to aggregate data from multiple Excel spreadsheets with the exact same structure into a single
final spreadsheet. A good use, for example, would be to combine multiple sources of identically structured spreadsheets that span different
time periods into a single spreadsheet in order to reduce the number of ETL jobs used to load the data into a database. The output excel
file will be placed in the directory above the named directory, and the output file name will be the name of the directory with the .xlsx
suffix added. For example, a folder named c:\path\to\TEST that contains 5 spreadsheets with identical structures will produce a file named
c:\path\to\TEST.xlsx
 
You may specify in a single optional parameter the number of the spreadsheet header row. During concatenation of the data from the various
spreadsheets, that row number will be assumed to be a header, and all rows before it will be skipped. The header row will appear only once
in the output file. Only the first worksheet in an excel workbook is processed. 

Warning: No worksheet structural validation is present. If your header rows do not match, the files will still be processed as-is and
data from the unmatching columns appended as-is to the final output. Output file is overwritten WITHOUT NOTICE.

## Compiling and Installation
This application should compile easily on Visual Studio 2015, with the Visual Studio Installer project type enabled. This requires a
separate download from Microsoft, that can be downloaded here:

https://marketplace.visualstudio.com/items?itemName=VisualStudioProductTeam.MicrosoftVisualStudio2015InstallerProjects

## Usage Examples
See Program.cs for detailed usage and command line parameters.

C:> excelMerge.exe c:\path\to\folder
C:> excelMerge.exe -h=2 c:\path\to\folder

## Version History
- 10/14/2016 - Initial unreleased development 
- 10/9/2017 - Initial release after bug fixes

## Requirements
- VC# / Visual Studio 2015
- .NET Framework 4.5.2
- Microsoft Excel installed on target machine for Microsoft.Office.Interop.Excel 

## Credits
by Jamie Anne Harrell

Various contributions and ideas for this project were borrowed from:
- Calabash Sisters        https://www.codeproject.com/Tips/715976/Solutions-to-Merge-Multiple-Excel-Worksheets-int (Solution 1 - CPOL license v1.02)
- Dmitry Martovoi         https://stackoverflow.com/questions/17367411/cannot-close-excel-exe-after-interop-process/17367570#17367570
- MSDN, of Course         https://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.workbooks.open.aspx

## License
Note: One file in this solution (MergeExcel.cs) is a derivative work of Calabash Sisters' Solution 1 referenced above, and distributed under the CPOL License v1.02. All other files:

Copyright (c) 2017 Goizueta Business School

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

<hr>

 ~=-=-=-=-=-=-=-=-=-=~  
~ Women Who Code Rock ~  
 ~=-=-=-=-=-=-=-=-=-=~