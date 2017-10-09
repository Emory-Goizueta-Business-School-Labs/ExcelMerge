using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;


namespace excelMerge
{

    public class MergeExcel
    {
        // The contents of this class is highly modified and thus should be considered "inspired" by the original work, which is licensed 
        // under the CPOL v1.02 license (https://www.codeproject.com/info/cpol10.aspx)
        //
        // Original Author: Calabash Sisters
        // Original Source: https://www.codeproject.com/Tips/715976/Solutions-to-Merge-Multiple-Excel-Worksheets-int
        //
        // TODO: Make sure to avoid the double-dot issue

        static string[] excelColumns =
        { "", // setup for 1-based indexing just for simplicity
        "A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z",
        "AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ",
        "BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW","BX","BY","BZ",
        "CA","CB","CC","CD","CE","CF","CG","CH","CI","CJ","CK","CL","CM","CN","CO","CP","CQ","CR","CS","CT","CU","CV","CW","CX","CY","CZ",
        "" // If you have more than 36 x 4 columns in your spreadsheet we're just gonna assume something broke somewhere else. 
        };


        Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
        Excel.Workbook bookDest = null;
        Excel.Worksheet sheetDest = null;

        //
        Excel.Workbook bookSource = null;
        Excel.Worksheet sheetSource = null;
        string[] _sourceFiles = null;
        string _destFile = string.Empty;
        //string _columnEnd = string.Empty;
        int _headerRowCount = 0;
        int _currentRowCount = 0;
        //int _columnCount = 0;

        public MergeExcel(string[] sourcefiles, string destFile, /* string columnEnd,*/ int headerRowCount)
        {
            bookDest = (Excel.Workbook)app.Workbooks.Add(Missing.Value);
            sheetDest = bookDest.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value) as Excel.Worksheet;
            sheetDest.Name = "data";
            _sourceFiles = sourcefiles;
            _destFile = destFile;
            //_columnEnd = columnEnd;
            _headerRowCount = headerRowCount;
        }

        // open worksheet
        void OpenBook(string fileName)
        {
            bookSource = app.Workbooks._Open(fileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            sheetSource = bookSource.Worksheets[1] as Excel.Worksheet;
        }

        void CloseBook()
        {
            bookSource.Close(false, Missing.Value, Missing.Value);
        }

        void CopyHeader()
        {
            string columnEnd = excelColumns[sheetSource.UsedRange.Columns.Count];
            Excel.Range range = sheetSource.get_Range("A1", columnEnd + _headerRowCount.ToString());
            range.Copy(sheetDest.get_Range("A1", Missing.Value));
            _currentRowCount += _headerRowCount;
        }

        void CopyData()
        {
            int sheetRowCount = sheetSource.UsedRange.Rows.Count;
            //int sheetColumnCount = sheetSource.UsedRange.Columns.Count;
            string columnEnd = excelColumns[sheetSource.UsedRange.Columns.Count];

            Excel.Range range = sheetSource.get_Range(string.Format("A{0}", _headerRowCount), columnEnd + sheetRowCount.ToString());
            range.Copy(sheetDest.get_Range(string.Format("A{0}", _currentRowCount), Missing.Value));
            _currentRowCount += range.Rows.Count;
        }

        void Save()
        {
            bookDest.Saved = true;
            bookDest.SaveCopyAs(_destFile);
        }

        void Quit()
        {
            app.Quit();
        }

        void DoMerge()
        {
            bool b = false;
            foreach (string strFile in _sourceFiles)
            {
                OpenBook(strFile);
                if (b == false)
                {
                    CopyHeader();
                    b = true;
                }
                CopyData();
                CloseBook();
            }
            Save();
            Quit();
        }

        public static void DoMerge(string[] sourceFiles, string destFile, /*string columnEnd,*/ int headerRowCount)
        {
            new MergeExcel(sourceFiles, destFile, /* columnEnd, */ headerRowCount).DoMerge();
        }

    }
}
