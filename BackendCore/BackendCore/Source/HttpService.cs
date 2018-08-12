using System;
using System.Net;
using System.Text;
using System.Threading;

namespace BackendCore.Source
{
    public class HttpService
    {
        private static string ServiceURL = "http://127.0.0.1:40000/";  // "http://localhost:40000/"

        private bool RunState = true;
        private ExcelParser.ExcelBase[] mExcelInfo = null;
   
        public HttpService(Object excel)
        {
            mExcelInfo = (ExcelParser.ExcelBase[])excel;
            // test_excel();

            HttpListener listener = new HttpListener();
            listener.Prefixes.Add(ServiceURL);
            listener.Start();

            System.Console.WriteLine("Http Listening..");

            while (RunState)
            {
                try
                {
                    HttpListenerContext context = listener.GetContext();
                    ThreadPool.QueueUserWorkItem(o => HandleRequest(context));
                }
                catch (Exception)
                {
                }
            }
        }

        private void test_excel()
        {
            Interface.JsonRequest request = new Interface.JsonRequest()
            {
                RequestID = 1,
                CarInfo = new Interface.JsonReq_CarInfo() {
                    Company = "현대자동차",
                    Model = "그랜저IG",
                    Trim = "가솔린 3.3 셀러브리티"
                },
                Cost = new Interface.JsonReq_Cost {
                    BasePrice = 41600000,
                    OptionPrice = 2600000,
                    OptionInfo = "HUD, 스마트센트II",
                    Deposit = 10,
                    PrePayment = 0
                },
                Commission = new Interface.JsonReq_Commission {
                    CMCommission = 2.5,
                    AGCommission = 5.5
                }
            };

            System.Console.WriteLine("<Request : JB Woori>");
            mExcelInfo[1].SetRequestInfo(request);
            var response1 = mExcelInfo[1].GetResonseInfo();
            PrintResponse(response1);

            System.Console.WriteLine("<Request : Hyosung>");
            mExcelInfo[2].SetRequestInfo(request);
            var response2 = mExcelInfo[2].GetResonseInfo();
            PrintResponse(response2);

            System.Console.WriteLine("<Request : Hana>");
            mExcelInfo[0].SetRequestInfo(request);
            var response0 = mExcelInfo[0].GetResonseInfo();
            PrintResponse(response0);
        }

        protected void PrintResponse(Interface.JsonResponseType resp)
        {
            System.Console.WriteLine("<Print Response Information>");
            System.Console.WriteLine("- CM Commission : {0}", resp.Commission.CMCommission);
            System.Console.WriteLine("- AG Commission : {0}", resp.Commission.AGCommission);

            System.Console.WriteLine("- 36개월 : {0} / {1} / {2}",
                resp.Payment.Fee36M.MonthlyFee, resp.Payment.Fee36M.AcquisitionPrice, resp.Payment.Fee36M.ResidualRate);
            System.Console.WriteLine("- 48개월 : {0} / {1} / {2}",
                resp.Payment.Fee48M.MonthlyFee, resp.Payment.Fee48M.AcquisitionPrice, resp.Payment.Fee48M.ResidualRate);
            System.Console.WriteLine("- 60개월 : {0} / {1} / {2}",
                resp.Payment.Fee60M.MonthlyFee, resp.Payment.Fee60M.AcquisitionPrice, resp.Payment.Fee60M.ResidualRate);
        }


        private string GetRequestData(HttpListenerRequest request)
        {
            byte[] readbuffer = new byte[1000];
            int len = request.InputStream.Read(readbuffer, 0, 1000);
            string jsonData = new string(Encoding.UTF8.GetChars(readbuffer, 0, len));

            return jsonData;
        }

        private void HandleRequest(HttpListenerContext context)
        {
            HttpListenerRequest request = context.Request;
            HttpListenerResponse response = context.Response;
            Interface.JsonRequest recvJson = null;

            try
            {
                string recvData = GetRequestData(request);
                recvJson = JsonParser.ParseJson(recvData);

                System.Console.WriteLine("<Receive Request>");
                System.Console.WriteLine("- URL : {0} / Method : {1}", request.Url.ToString(), request.HttpMethod);
                System.Console.WriteLine("- jsonData : {0}", recvData);
                System.Console.WriteLine("<Parsing Request Data>");
                System.Console.WriteLine("- Model : {0} /  Base Cost : {1}", recvJson.CarInfo.Model, recvJson.Cost.BasePrice);
                System.Console.WriteLine("-------------------------------------------------");

                if (recvJson.Cost.BasePrice == -1) RunState = false;
            }
            catch (Exception)
            {
                System.Console.WriteLine("HandleRequest : Request Error! \n");
                // Client disconnected or some other error - ignored for this example
            }

            try
            {
                if (recvJson != null)
                {
                    Interface.JsonResponse sendJson = CalculateFee(recvJson);
                    ResponseToRequester(response, sendJson);
                }
            }
            catch (Exception)
            {
                System.Console.WriteLine("HandleRequest : make Resonse Error! \n");
                // Client disconnected or some other error - ignored for this example
                ResponseToRequesterError(response);
            }
        }

        private Interface.JsonResponse CalculateFee(Interface.JsonRequest request)
        {
            Interface.JsonResponse response = new Interface.JsonResponse();
            response.RequestID = request.RequestID;
            response.Response = new Interface.JsonResponseType[mExcelInfo.Length];
            // JsonSend[] sendData = new JsonSend[mExcelInfo.Length];

            for (int i = 0; i < mExcelInfo.Length; i++)
            {
                System.Console.WriteLine("CalculateFee / COM : {0}", mExcelInfo[i].GetCapitalName());
                System.Console.WriteLine("<Request : {0} >", mExcelInfo[i].GetCapitalName());

                mExcelInfo[i].SetRequestInfo(request);
                response.Response[i] = mExcelInfo[i].GetResonseInfo();
                PrintResponse(response.Response[i]);

/*
                mExcelInfo[i].SetPrice(parseData.Price);
                mExcelInfo[i].SetRate(parseData.Rate);

                string com = mExcelInfo[i].GetCapitalName();
                int rate = mExcelInfo[i].GetRate(parseData.Rate);
                int m36 = mExcelInfo[i].GetFeeM36();
                int m48 = mExcelInfo[i].GetFeeM48();
                int m60 = mExcelInfo[i].GetFeeM60();

                sendData[i] = new JsonSend { Com = com, Rate = rate, FeeM36 = m36, FeeM48 = m48, FeeM60 = m60 };
*/
            }

            return response;
        }

        private void ResponseToRequester(HttpListenerResponse response, Interface.JsonResponse calculateData)
        {
            try
            {
                // Send Message
                string responseString = JsonParser.ComposeJson(calculateData);
                System.Console.WriteLine(" -- Json String : {0}", responseString);
                byte[] buffer = Encoding.UTF8.GetBytes(responseString);

                response.StatusCode = 200;
                response.SendChunked = true;
                response.OutputStream.Write(buffer, 0, buffer.Length);
                response.OutputStream.Close();
            }
            catch (Exception)
            {
                // Client disconnected or some other error - ignored for this example
            }
        }

        private void ResponseToRequesterError(HttpListenerResponse response)
        {
            try
            {
                // Send Message
                string responseString = "Internal Error";
                byte[] buffer = Encoding.UTF8.GetBytes(responseString);

                response.StatusCode = 202;
                response.SendChunked = true;
                response.OutputStream.Write(buffer, 0, buffer.Length);
                response.OutputStream.Close();
            }
            catch (Exception)
            {
                // Client disconnected or some other error - ignored for this example
            }
        }
    }
}
