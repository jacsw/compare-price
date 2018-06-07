using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace BackendCore.Source.ExcelParser
{
    class ExcelPosition
    {
        int Row;
        int Col;
    }

    abstract class ExcelBase
    {
        //
        // ExcelBase Common 

        public ExcelBase(Config.CapitalData pConfig)
        {
            mConfig = pConfig;

            mCapitalName = pConfig.CapitalName;
            mExcelFileName = pConfig.Config.ExcelFile;
            mSheetName = pConfig.Config.Worksheet;
        }

        public string GetCapitalName()
        {
            return mCapitalName;
        }

        public void Init()
        {
            try
            {
                System.Console.Write("Data Loading...{0}({1})...", mCapitalName, mConfig.Config.ExcelFile);
                mWorkbook = ExcelInstance.getInstance().Workbooks.Open(mExcelFileName);
                mWorksheet = mWorkbook.Sheets[mSheetName];
                System.Console.WriteLine("Completed");

                SetPositionfromConfig();
            }
            finally
            {
            }
        }

        public void Deinit()
        {
            try
            {
                if (mWorkbook != null) mWorkbook.Close(false);
            }
            finally
            {
                ExcelInstance.ReleaseExcelObject(mWorksheet);
                ExcelInstance.ReleaseExcelObject(mWorkbook);
            }
        }

        protected Config.CapitalData mConfig;
        protected Excel.Workbook mWorkbook = null;
        protected Excel.Worksheet mWorksheet = null;

        protected string mCapitalName;
        protected string mExcelFileName;
        protected string mSheetName;

        // 
        // 
        //ExcelBase SetPosition
        protected void SetPositionfromConfig()
        {
            mPosCarInfo = mConfig.Config.CarInfo;
            mPosPrice = mConfig.Config.Price;
            mPosPayment = mConfig.Config.Payment;
      
            SetPositionfromConfigEach();
        }

        protected abstract void SetPositionfromConfigEach();

        Config.JsonConfig_CarInfo mPosCarInfo;
        Config.JsonConfig_Price mPosPrice;
        Config.JsonConfig_Payment mPosPayment;

        public abstract void SetRequestInfo(Interface.JsonRequest request);

        public abstract Interface.JsonResponseType GetResonseInfo();
    }
}
