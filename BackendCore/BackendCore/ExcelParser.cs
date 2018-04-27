using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Xml;

namespace ExcelTest
{
    class ExcelParser
    {
        public static void ReleaseExcelObject(object obj)
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


        static Excel.Application sExcelApp = new Excel.Application();
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

        public string getCompanyName() { return mCompanyName;  }

        private string mCompanyName;
        private string mExcelFileName;
        private string mSheetName;
 
        private Excel.Workbook mWorkbook = null;
        private Excel.Worksheet mWorksheet = null;

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

        int mGtotRow = 0, mGtotCol = 0;
        int mRentRow = 0, mRentCol = 0;

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
        
    }
}


// ws2.Cells[14, 59].Value = 40000000 + i * 10000000;

// System.Console.WriteLine("{0,10} {1,10} {2,10}",
// mWorksheet.Cells[16, 50].Value,
//  mWorksheet.Cells[27, 9].Value, ws2.Cells[30, 8].Value);
