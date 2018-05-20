using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace BackendCore
{
    class ExcelParser
    {
        public ExcelParser(string _company, string _filename, string _sheetname)
        {
            mCompanyName = _company;
            mExcelFileName = _filename;
            mSheetName = _sheetname;
        }

        static public void ModuleInit()
        {
            sExcelApp = new Excel.Application();
        }

        static public void ModuleDeInit()
        {
            sExcelApp.Quit();
            ReleaseExcelObject(sExcelApp);
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

        public void SetPricePos(int row, int col)
        {
            mPriceRow = row;
            mPriceCol = col;
        }

        public void SetPrice(int price)
        {
            if (mPriceRow == 0) return;

            try
            {
                mWorksheet.Cells[mPriceRow, mPriceCol].Value = price;
            }
            catch (Exception e)
            {
                System.Console.WriteLine("- Calculate Error Fee M36 \n{0}", e.StackTrace);
            }
        }

        public void SetRatePos(int row, int col)
        {
            mRateRow = row;
            mRateCol = col;
        }

        public void SetRate(int price)
        {
            if (mRateRow == 0) return;

            try
            {
                mWorksheet.Cells[mRateRow, mRateCol].Value = price;
            }
            catch (Exception e)
            {
                System.Console.WriteLine("- Calculate Error Fee M36 \n{0}", e.StackTrace);
            }
        }

        public int GetRate(int rate)
        {
            int ret = 0;
            if (mRateRow == 0) return rate;

            try
            {
                var cell = mWorksheet.Cells[mRateRow, mRateCol];
                if (mWorksheet.Cells[mRateRow, mRateCol].Value.GetType() == typeof(string))
                {
                    Int32.TryParse(mWorksheet.Cells[mRateRow, mRateCol].Value, out ret);
                }
                else if (mWorksheet.Cells[mRateRow, mRateCol].Value.GetType() == typeof(double))
                {
                    ret = Convert.ToInt32(mWorksheet.Cells[mRateRow, mRateCol].Value);
                }
            }
            catch (Exception e)
            {
                System.Console.WriteLine("- Calculate Error Fee M36 \n{0}", e.StackTrace);
            }

            return ret;

        }

        public void SetFeeM36Pos(int row, int col)
        {
            mFeeM36Row = row;
            mFeeM36Col = col;
        }
 
        public int GetFeeM36()
        {
            int ret = 0;
            if (mFeeM36Row == 0) return 0;

            try
            {
                if (mWorksheet.Cells[mFeeM36Row, mFeeM36Col].Value.GetType() == typeof(string))
                {
                    Int32.TryParse(mWorksheet.Cells[mFeeM36Row, mFeeM36Col].Value, out ret);
                }
                else if (mWorksheet.Cells[mFeeM36Row, mFeeM36Col].Value.GetType() == typeof(double))
                {
                    ret = Convert.ToInt32(mWorksheet.Cells[mFeeM36Row, mFeeM36Col].Value);
                }
            }
            catch(Exception e)
            {
                System.Console.WriteLine("- Calculate Error Fee M36 \n{0}", e.StackTrace);
            }

            return ret;
        }

        public void SetFeeM48Pos(int row, int col)
        {
            mFeeM48Row = row;
            mFeeM48Col = col;
        }

        public int GetFeeM48()
        {
            int ret = 0;
            if (mFeeM48Row == 0) return 0;

            try
            {
                if (mWorksheet.Cells[mFeeM48Row, mFeeM48Col].Value.GetType() == typeof(string))
                {
                    Int32.TryParse(mWorksheet.Cells[mFeeM48Row, mFeeM48Col].Value, out ret);
                }
                else if (mWorksheet.Cells[mFeeM48Row, mFeeM48Col].Value.GetType() == typeof(double))
                {
                    ret = Convert.ToInt32(mWorksheet.Cells[mFeeM48Row, mFeeM48Col].Value);
                }
            }
            catch (Exception e)
            {
                System.Console.WriteLine("- Calculate Error Fee M48 \n{0}", e.StackTrace);
            }

            return ret;
        }

        public void SetFeeM60Pos(int row, int col)
        {
            mFeeM60Row = row;
            mFeeM60Col = col;
        }

        public int GetFeeM60()
        {
            int ret = 0;
            if (mFeeM60Row == 0) return 0;

            try
            {
                if (mWorksheet.Cells[mFeeM60Row, mFeeM60Col].Value.GetType() == typeof(string))
                {
                    Int32.TryParse(mWorksheet.Cells[mFeeM60Row, mFeeM60Col].Value, out ret);
                }
                else if (mWorksheet.Cells[mFeeM60Row, mFeeM60Col].Value.GetType() == typeof(double))
                {
                    ret = Convert.ToInt32(mWorksheet.Cells[mFeeM60Row, mFeeM60Col].Value);
                }
            }
            catch (Exception e)
            {
                System.Console.WriteLine("- Calculate Error Fee M60 \n{0}", e.StackTrace);
            }

            return ret;
        }

        private Excel.Workbook mWorkbook = null;
        private Excel.Worksheet mWorksheet = null;

        private string mCompanyName;
        private string mExcelFileName;
        private string mSheetName;

        private int mPriceRow = 0, mPriceCol = 0;
        private int mRateRow = 0, mRateCol = 0;

        private int mFeeM36Row = 0, mFeeM36Col = 0;
        private int mFeeM48Row = 0, mFeeM48Col = 0;
        private int mFeeM60Row = 0, mFeeM60Col = 0;

        private static Excel.Application sExcelApp = new Excel.Application();

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
