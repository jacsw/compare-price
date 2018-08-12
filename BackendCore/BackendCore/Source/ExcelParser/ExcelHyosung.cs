using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BackendCore.Source.Interface;

namespace BackendCore.Source.ExcelParser
{
    class ExcelHyosung : ExcelBase
    {
        Dictionary<string, int> mComMapID;
        public ExcelHyosung(Config.CapitalData pConfig) : base(pConfig)
        {
            mComMapID = new Dictionary<string, int>();
            mComMapID.Add("현대자동차", 1);
            mComMapID.Add("기아자동차", 2);
            mComMapID.Add("GM쉐보레", 4);
            mComMapID.Add("르노삼성자동차", 3);
            mComMapID.Add("쌍용자동차", 5);
        }

        public override JsonResponseType GetResonseInfo()
        {
            var resp = MakeResponse();

            resp.CapitalName = "효성캐피탈";

            var CarInfo = mConfig.Config.CarInfo;
            //resp.CarInfo.Company = mWorksheet.Cells[CarInfo.Company.Row, CarInfo.Company.Col].Value;
            resp.CarInfo.Model = mWorksheet.Cells[CarInfo.Model.Row, CarInfo.Model.Col].Value;
            resp.CarInfo.Trim = "";

            var Extra = mConfig.Config.Extra;
            resp.Commission.CMCommission = mWorksheet.Cells[Extra[1].Row, Extra[1].Col].Value;
            resp.Commission.AGCommission = mWorksheet.Cells[Extra[2].Row, Extra[2].Col].Value;

            GetCommonResponse(resp);

            return resp;
        }

        public override void SetRequestInfo(JsonRequest request)
        {
            int comnameid = mComMapID[request.CarInfo.Company];
            string comname = mComMap[request.CarInfo.Company];
            string model = comname + "/" + request.CarInfo.Model + "/" + request.CarInfo.Trim;
            string capitalname = CarList.getInstance().GetCarName(model).Hyosung;

            var CarInfo = mConfig.Config.CarInfo;
            mWorksheet.Cells[CarInfo.Company.Row, CarInfo.Company.Col] = comnameid;
            mWorksheet.Cells[CarInfo.Model.Row, CarInfo.Model.Col] = capitalname;

            var Price = mConfig.Config.Price;
            mWorksheet.Cells[Price.BasePrice.Row, Price.BasePrice.Col] = request.Cost.BasePrice;
            mWorksheet.Cells[Price.OptionPrice.Row, Price.OptionPrice.Col] = request.Cost.OptionPrice;
            mWorksheet.Cells[Price.OptionInfo.Row, Price.OptionInfo.Col] = request.Cost.OptionInfo;

            var Payment = mConfig.Config.Payment;
            mWorksheet.Cells[Payment.Deposit.Row, Payment.Deposit.Col] = request.Cost.Deposit.ToString() + "%";
            mWorksheet.Cells[Payment.PrePayment.Row, Payment.PrePayment.Col] = request.Cost.PrePayment.ToString() + "%";

            var Extra = mConfig.Config.Extra;

            mWorksheet.Cells[Extra[1].Row, Extra[1].Col] = request.Commission.CMCommission.ToString() + "%";
            mWorksheet.Cells[Extra[2].Row, Extra[2].Col] = request.Commission.AGCommission.ToString() + "%";
            mWorksheet.Cells[Extra[3].Row, Extra[3].Col] = 4;  // "정비 Basic"
            mWorksheet.Cells[Extra[4].Row, Extra[4].Col] = 2;  // "Snow Tire : 미제공"

            System.Console.WriteLine("<Receive RequestSetRequestInfo>");
            System.Console.WriteLine("- Model  : {0} / Capital Name : {1}", model, capitalname);
            System.Console.WriteLine("- 배기량 : {0}", mWorksheet.Cells[14, 65].Value);
            System.Console.WriteLine("- 총  액 : {0}", mWorksheet.Cells[Price.TotalPrice.Row, Price.TotalPrice.Col].Value);
            System.Console.WriteLine("- 옵  션 : {0}", mWorksheet.Cells[Price.OptionInfo.Row, Price.OptionInfo.Col].Value);
        }

        protected override void SetPositionfromConfigEach()
        {
            // throw new NotImplementedException();
        }
    }
}
