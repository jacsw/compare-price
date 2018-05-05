using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace BackendCore
{
    class ExcelParser
    {
        private static Excel.Application sExcelApp = new Excel.Application();

        static public void ModuleInit()
        {
            sExcelApp = new Excel.Application();
        }

        static public void ModuleDeInit()
        {
            sExcelApp.Quit();
            ReleaseExcelObject(sExcelApp);
        }

        public ExcelParser(string _company, string _filename, string _sheetname)
        {
            mCompanyName = _company;
            mExcelFileName = _filename;
            mSheetName = _sheetname;
        }

        public string GetCompanyName()
        {
            return mCompanyName;
        }

        public void Init()
        {
            try
            {
                System.Console.Write("Data Loading...{0}({1})...", mCompanyName, mExcelFileName);
                mWorkbook = sExcelApp.Workbooks.Open(mExcelFileName);
                mWorksheet = mWorkbook.Sheets[mSheetName];
                System.Console.WriteLine("Completed");
            }
            finally
            {
            }
        }

        public void Deinit()
        {
            try
            {
                mWorkbook.Close(false);
            }
            finally
            {
                ReleaseExcelObject(mWorksheet);
                ReleaseExcelObject(mWorkbook);
            }
        }

        public void SetGTotalPos(int row, int col)
        {
            mGtotRow = row;
            mGtotCol = col;
        }

        public void SetRentCostPos(int row, int col)
        {
            mRentRow = row;
            mRentCol = col;
        }

        public int GetRentCost(int carValue)
        {
            int ret = 0;
            try
            {
                mWorksheet.Cells[mGtotRow, mGtotCol].Value = carValue;
                if (mWorksheet.Cells[mRentRow, mRentCol].Value.GetType() == typeof(string))
                {
                    Int32.TryParse(mWorksheet.Cells[mRentRow, mRentCol].Value, out ret);
                }
                else if (mWorksheet.Cells[mRentRow, mRentCol].Value.GetType() == typeof(double))
                {
                    ret = Convert.ToInt32(mWorksheet.Cells[mRentRow, mRentCol].Value);
                }
            }
            finally
            {
            }
            return ret;
        }

        private Excel.Workbook mWorkbook = null;
        private Excel.Worksheet mWorksheet = null;

        private string mCompanyName;
        private string mExcelFileName;
        private string mSheetName;

        private int mGtotRow = 0, mGtotCol = 0;
        private int mRentRow = 0, mRentCol = 0;
  
        private static void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

    }
}
