/*
 * Program: excelMerge.exe
 * Author: Jamie Anne Harrell (jamie.harrell@emory.edu)
 * Date: 10/14/2016
 * 
 * 
 * NAME
 * excelMerge.exe - command line utility to merge all excel documents within a directory into a single excel workbook.
 * 
 * SYNOPSIS
 * excelMerge.exe [OPTION] [source_directory]
 * 
 * DESCRIPTION
 * excelMerge.exe has a very specific use case - to aggregate data from multiple Excel spreadsheets with the exact same structure into a single
 * final spreadsheet. A good use, for example, would be to combine multiple sources of identically structured spreadsheets that span different
 * time periods into a single spreadsheet in order to reduce the number of ETL jobs used to load the data into a database. The output excel
 * file will be placed in the directory above the named directory, and the output file name will be the name of the directory with the .xlsx
 * suffix added. For example, a folder named c:\path\to\TEST that contains 5 spreadsheets with identical structures will produce a file named
 * c:\path\to\TEST.xlsx
 * 
 * You may specify in a single optional parameter the number of the spreadsheet header row. During concatenation of the data from the various
 * spreadsheets, that row number will be assumed to be a header, and all rows before it will be skipped. The header row will appear only once
 * in the output file. Only the first worksheet in an excel workbook is processed. Non-excel files (not xls or xlsx) in the direcectory are
 * ignored.
 * 
 * *** Warning: no worksheet structural validation is present. If your header rows do not match, the files will still be processed as-is and
 * *** data from the unmatching columns appended as-is to the final output. Output file is overwritten WITHOUT NOTICE.
 * 
 * * TODO: MAKE THIS MORE FLEXIBLE AND ADD HEADER VALIDATION
 * 
 * GENERAL
 * Only a single operational mode is supported with one optional parameter as noted below.
 * 
 * excelMerge.exe [PARAMETER] [source_directory]
 * 
 * PARAMETERS
 *  -h=N, --header=N, /v=N, /header=N   where N is the integer row number where the header is in the spreadsheet. Rows prior to this are skipped.
 *  
 * EXAMPLES:
 *  
 *  excelMerge.exe c:\path\to\folder
 *  
 *  excelMerge.exe -h=2 c:\path\to\folder
 *  
 *  
 * CREDITS: Various contributions and ideas for this project were borrowed from:
 *  Calabash Sisters        https://www.codeproject.com/Tips/715976/Solutions-to-Merge-Multiple-Excel-Worksheets-int (Solution 1 - CPOL license v1.02)
 *  Dmitry Martovoi         https://stackoverflow.com/questions/17367411/cannot-close-excel-exe-after-interop-process/17367570#17367570
 *  MSDN, of Course         https://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.workbooks.open.aspx
 * 
 * 
 * License Notes: the class MergeExcel.cs is licensed under the Open Source CPOL v1.02 license https://www.codeproject.com/info/cpol10.aspx
 *
 * All other files:
 *  
 * License: MIT / Open Source
 * 
 * Copyright (c) 2017 Emory Goizueta Business School
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE. 
 * 
 */

using System;
using System.Linq;
using System.IO;

namespace excelMerge
{


    class Program
    {

        public static int headerRow = 1;

        static void Usage()
        {
            Console.WriteLine("EXCELMERGE USAGE:");
            Console.WriteLine("========================================");
            Console.WriteLine("excelMerge.exe [PARAMETER] [source_directory]");
            Console.WriteLine(" PARAMETER:");
            Console.WriteLine(" -h=N, --header=N, /v=N, /header=N");
            Console.WriteLine(" EXAMPLES:");
            Console.WriteLine(@" excelMerge.exe c:\path\to\folder");
            Console.WriteLine(@" excelMerge.exe -h=2 c:\path\to\folder");
        }

        static int Main(string[] args)
        {
            int headerRow = 1;
            string headerRowOption = "";
            String directory = "";
            switch (args.Length) {
                case 1:
                    directory = args[0];
                    break;
                case 2:
                    headerRowOption = args[0];
                    directory = args[1];
                    char[] separator = { '=' }; 
                    string[] tokens = headerRowOption.Split(separator, 2);
                    if (tokens.Count() != 2) { Usage(); return 1; }
                    if ((tokens[0] != "-h") && (tokens[0] != "--header") && (tokens[0] != "/h") && (tokens[0] != "/header") ){ Usage(); return 2; }
                    if (!int.TryParse(tokens[1],out headerRow)) { Usage(); return 3; }
                    break;
                default:
                    Usage();
                    return (1);
            }
               
            // make directory relative if not found absolutely
            if (!Directory.Exists(directory))
            {
                directory = System.IO.Directory.GetCurrentDirectory() + @"\" + directory;
                if (!Directory.Exists(directory))
                {
                   System.Console.WriteLine("Invalid directory");
                   return (4);
                }
            }

            // Get list of all files in the directory
            string[] fileEntries = Directory.GetFiles(directory);
            string destination = directory + ".xlsx";
            System.Console.WriteLine("Processing directory " + directory + " into " + destination);
            try
            {
                MergeExcel.DoMerge(fileEntries, destination, headerRow);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error processing files: '{0}'", e.Message);
                return (-1);
            }
            
            return (0);
        }
    }


   
}