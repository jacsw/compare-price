using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

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

        protected Dictionary<string, string> mComMap;

        public ExcelBase(Config.CapitalData pConfig)
        {
            mConfig = pConfig;

            mCapitalName = pConfig.CapitalName;
            mExcelFileName = pConfig.Config.ExcelFile;
            mSheetName = pConfig.Config.Worksheet;

            mComMap = new Dictionary<string, string>();
            mComMap.Add("현대자동차", "현대");
            mComMap.Add("기아자동차", "기아");
            mComMap.Add("GM쉐보레", "GM");
            mComMap.Add("르노삼성자동차", "삼성");
            mComMap.Add("쌍용자동차", "쌍용");
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

        protected Interface.JsonResponseType MakeResponse()
        {
            Interface.JsonResponseType resp = new Interface.JsonResponseType();

            resp.CapitalName = mCapitalName;
            resp.CarInfo     = new Interface.JsonResp_CarInfo();
            resp.Commission  = new Interface.JsonResp_Commission();
            resp.Insurance   = new Interface.JsonResp_Insurance();
            resp.Shipment    = new Interface.JsonResp_Shipment();

            resp.Payment     = new Interface.JsonResp_Payment();
            resp.Payment.Fee36M = new Interface.JsonResp_Payment_Fee();
            resp.Payment.Fee48M = new Interface.JsonResp_Payment_Fee();
            resp.Payment.Fee60M = new Interface.JsonResp_Payment_Fee();

            return resp;
        }

        protected void GetCommonResponse(Interface.JsonResponseType resp)
        {
            var Payment = mConfig.Config.Payment;
            mWorksheet.Cells[Payment.Duration.Row, Payment.Duration.Col] = "36";
            
            resp.Payment.Fee36M.MonthlyFee = (int)(mWorksheet.Cells[Payment.MonthlyFee.Row, Payment.MonthlyFee.Col].Value);
            resp.Payment.Fee36M.AcquisitionPrice = (int)(mWorksheet.Cells[Payment.ResidualValue.Row, Payment.ResidualValue.Col].Value);
            resp.Payment.Fee36M.ResidualRate = mWorksheet.Cells[Payment.ResidualRate.Row, Payment.ResidualRate.Col].Value;

            mWorksheet.Cells[Payment.Duration.Row, Payment.Duration.Col] = "48";
            resp.Payment.Fee48M.MonthlyFee = (int)(mWorksheet.Cells[Payment.MonthlyFee.Row, Payment.MonthlyFee.Col].Value);
            resp.Payment.Fee48M.AcquisitionPrice = (int)(mWorksheet.Cells[Payment.ResidualValue.Row, Payment.ResidualValue.Col].Value);
            resp.Payment.Fee48M.ResidualRate = mWorksheet.Cells[Payment.ResidualRate.Row, Payment.ResidualRate.Col].Value;

            mWorksheet.Cells[Payment.Duration.Row, Payment.Duration.Col] = "60";
            resp.Payment.Fee60M.MonthlyFee = (int)(mWorksheet.Cells[Payment.MonthlyFee.Row, Payment.MonthlyFee.Col].Value);
            resp.Payment.Fee60M.AcquisitionPrice = (int)(mWorksheet.Cells[Payment.ResidualValue.Row, Payment.ResidualValue.Col].Value);
            resp.Payment.Fee60M.ResidualRate = mWorksheet.Cells[Payment.ResidualRate.Row, Payment.ResidualRate.Col].Value;
        }
    }
}
