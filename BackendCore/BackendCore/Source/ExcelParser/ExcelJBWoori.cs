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
        public ExcelJBWoori(Config.CapitalData pConfig) : base(pConfig)
        {
        }

        public override JsonResponseType GetResonseInfo()
        {
            var resp = MakeResponse();

            resp.CapitalName = "JB우리캐피탈";

            var CarInfo = mConfig.Config.CarInfo;
            resp.CarInfo.Company = mWorksheet.Cells[CarInfo.Company.Row, CarInfo.Company.Col].Value;
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
            mWorksheet.Cells[Payment.Deposit.Row, Payment.Deposit.Col] = request.Cost.Deposit / 100.0;
            mWorksheet.Cells[Payment.PrePayment.Row, Payment.PrePayment.Col] = request.Cost.PrePayment;

            var Extra = mConfig.Config.Extra;
            mWorksheet.Cells[Extra[1].Row, Extra[1].Col] = request.Commission.CMCommission.ToString() + "%";
            mWorksheet.Cells[Extra[2].Row, Extra[2].Col] = request.Commission.AGCommission.ToString() + "%";
            mWorksheet.Cells[Extra[3].Row, Extra[3].Col] = "정비제외";
            mWorksheet.Cells[Extra[4].Row, Extra[4].Col] = "셀프";

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
