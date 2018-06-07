using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BackendCore.Source.Interface;

namespace BackendCore.Source.ExcelParser
{
    class ExcelJBWoori : ExcelBase
    {
        Dictionary<string, string> mComMap;
        public ExcelJBWoori(Config.CapitalData pConfig) : base(pConfig)
        {
            mComMap = new Dictionary<string, string>();
            mComMap.Add("현대자동차", "현대");
            mComMap.Add("기아자동차", "기아");
        }

        public override JsonResponseType GetResonseInfo()
        {
            JsonResponseType resp = new JsonResponseType();

            resp.CapitalName = mCapitalName;
            resp.CarInfo = new JsonResp_CarInfo();
            resp.Commission = new JsonResp_Commission();
            resp.Insurance = new JsonResp_Insurance();
            resp.Shipment = new JsonResp_Shipment();
            resp.Payment = new JsonResp_Payment();
            resp.Payment.Fee36M = new JsonResp_Payment_Fee();
            resp.Payment.Fee48M = new JsonResp_Payment_Fee();
            resp.Payment.Fee60M = new JsonResp_Payment_Fee();


            var CarInfo = mConfig.Config.CarInfo;
            resp.CarInfo.Company = mWorksheet.Cells[CarInfo.Company.Row, CarInfo.Company.Col].Value;
            resp.CarInfo.Model = mWorksheet.Cells[CarInfo.Model.Row, CarInfo.Model.Col].Value;
            resp.CarInfo.Trim = "";

            var Extra = mConfig.Config.Extra;
            resp.Commission.CMCommission = mWorksheet.Cells[Extra[1].Row, Extra[1].Col].Value;
            resp.Commission.AGCommission = mWorksheet.Cells[Extra[2].Row, Extra[2].Col].Value;

            var Payment = mConfig.Config.Payment;
            mWorksheet.Cells[Payment.Duration.Row, Payment.Duration.Col] = 36;
            resp.Payment.Fee36M.MonthlyFee = (int)( mWorksheet.Cells[Payment.MonthlyFee.Row, Payment.MonthlyFee.Col].Value );
            resp.Payment.Fee36M.AcquisitionPrice = (int)( mWorksheet.Cells[Payment.ResidualValue.Row, Payment.ResidualValue.Col].Value );
            resp.Payment.Fee36M.ResidualRate = mWorksheet.Cells[Payment.ResidualRate.Row, Payment.ResidualRate.Col].Value;

            mWorksheet.Cells[Payment.Duration.Row, Payment.Duration.Col] = 48;
            resp.Payment.Fee48M.MonthlyFee = (int)(mWorksheet.Cells[Payment.MonthlyFee.Row, Payment.MonthlyFee.Col].Value);
            resp.Payment.Fee48M.AcquisitionPrice = (int)(mWorksheet.Cells[Payment.ResidualValue.Row, Payment.ResidualValue.Col].Value);
            resp.Payment.Fee48M.ResidualRate = mWorksheet.Cells[Payment.ResidualRate.Row, Payment.ResidualRate.Col].Value;

            mWorksheet.Cells[Payment.Duration.Row, Payment.Duration.Col] = 60;
            resp.Payment.Fee60M.MonthlyFee = (int)(mWorksheet.Cells[Payment.MonthlyFee.Row, Payment.MonthlyFee.Col].Value);
            resp.Payment.Fee60M.AcquisitionPrice = (int)(mWorksheet.Cells[Payment.ResidualValue.Row, Payment.ResidualValue.Col].Value);
            resp.Payment.Fee60M.ResidualRate = mWorksheet.Cells[Payment.ResidualRate.Row, Payment.ResidualRate.Col].Value;





            System.Console.WriteLine("- CM Commission : {0}", mWorksheet.Cells[Extra[1].Row, Extra[1].Col].Value);
            System.Console.WriteLine("- AG Commission : {0}", mWorksheet.Cells[Extra[2].Row, Extra[2].Col].Value);

            System.Console.WriteLine("- 36개월 : {0} / {1} / {2}",
                resp.Payment.Fee36M.MonthlyFee, resp.Payment.Fee36M.AcquisitionPrice, resp.Payment.Fee36M.ResidualRate);
            System.Console.WriteLine("- 48개월 : {0} / {1} / {2}",
                resp.Payment.Fee48M.MonthlyFee, resp.Payment.Fee48M.AcquisitionPrice, resp.Payment.Fee48M.ResidualRate);
            System.Console.WriteLine("- 60개월 : {0} / {1} / {2}",
                resp.Payment.Fee60M.MonthlyFee, resp.Payment.Fee60M.AcquisitionPrice, resp.Payment.Fee60M.ResidualRate);


            return resp;
        }

        public override void SetRequestInfo(JsonRequest request)
        {
            string comname = mComMap[request.CarInfo.Company];
            string model = comname + "/" + request.CarInfo.Model + "/" + request.CarInfo.Trim;
            string capitalname = CarList.getInstance().GetCarName(model).JBWoori;

            var CarInfo = mConfig.Config.CarInfo;
            mWorksheet.Cells[CarInfo.Company.Row, CarInfo.Company.Col] = comname;
            mWorksheet.Cells[CarInfo.Model.Row, CarInfo.Model.Col] = capitalname;

            var Price = mConfig.Config.Price;
            mWorksheet.Cells[Price.BasePrice.Row, Price.BasePrice.Col] = request.Cost.BasePrice;
            mWorksheet.Cells[Price.OptionPrice.Row, Price.OptionPrice.Col] = request.Cost.OptionPrice;
            mWorksheet.Cells[Price.OptionInfo.Row, Price.OptionInfo.Col] = request.Cost.OptionInfo;

            var Payment = mConfig.Config.Payment;
            // mWorksheet.Cells[Payment.Deposit.Row, Payment.Deposit.Col] = request.Cost.Deposit;
            mWorksheet.Cells[Payment.PrePayment.Row, Payment.PrePayment.Col] = request.Cost.PrePayment;

            var Extra = mConfig.Config.Extra;
            mWorksheet.Cells[Extra[1].Row, Extra[1].Col] = request.Commission.CMCommission.ToString() + "%";
            mWorksheet.Cells[Extra[2].Row, Extra[2].Col] = request.Commission.AGCommission.ToString() + "%";



            System.Console.WriteLine("<Receive RequestSetRequestInfo>");
            System.Console.WriteLine("- Model  : {0} / Capital Name : {1}", model, capitalname);
            System.Console.WriteLine("- 배기량 : {0}", mWorksheet.Cells[14, 55].Value);
            System.Console.WriteLine("- 총  액 : {0}", mWorksheet.Cells[Price.TotalPrice.Row, Price.TotalPrice.Col].Value);
            System.Console.WriteLine("- 옵  션 : {0}", mWorksheet.Cells[Price.OptionInfo.Row, Price.OptionInfo.Col].Value);
        }

        protected override void SetPositionfromConfigEach()
        {
            // throw new NotImplementedException();
        }
    }
}
