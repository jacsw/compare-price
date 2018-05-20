// Interface 
// * Name    : JsonResponse
// * Version : 0.5
// * Usage   : Response Format from back end to front end

using System.Runtime.Serialization;

namespace BackendCore.Source.Interface
{
    //
    // Main Type : Response
    [DataContract]
    public class JsonResponse
    {
        [DataMember] public int RequestID;
        [DataMember] public JsonResponseType[] Response;
    }

    //
    // define elements of JsonResponse.
    [DataContract]
    public class JsonResponseType
    {
        [DataMember] public string CapitalName;
        [DataMember] public JsonResp_CarInfo CarInfo;
        [DataMember] public JsonResp_Payment Payment;
        [DataMember] public JsonResp_Commission Commission;
        [DataMember] public JsonResp_Insurance Insurance;
        [DataMember] public JsonResp_Shipment Shipment;
    }

    [DataContract]
    public class JsonResp_CarInfo
    {
        [DataMember] public string Company;
        [DataMember] public string Model;
        [DataMember] public string Trim;
    }

    [DataContract]
    public class JsonResp_Payment_Fee
    {
        [DataMember] public int MonthlyFee;
        [DataMember] public int AcquisitionPrice;
        [DataMember] public int ResidualValue;
        [DataMember] public int ResidualRate;
    }

    [DataContract]
    public class JsonResp_Payment
    {
        [DataMember] public int Deposit;
        [DataMember] public int PrePayment;
        [DataMember] public JsonResp_Payment_Fee Fee36M;
        [DataMember] public JsonResp_Payment_Fee Fee48M;
        [DataMember] public JsonResp_Payment_Fee Fee60M;
    }

    [DataContract]
    public class JsonResp_Commission
    {
        [DataMember] public double CMCommission;
        [DataMember] public double AGCommission;
    }

    [DataContract]
    public class JsonResp_Insurance
    {
        [DataMember] public string CorporateType;
        [DataMember] public int DriverAge;
        [DataMember] public int Coverage1;
        [DataMember] public int Coverage2;
        [DataMember] public int Exemption;
    }

    [DataContract]
    public class JsonResp_Shipment
    {
        [DataMember] public string SalesType;
        [DataMember] public string Discount;
        [DataMember] public string Delivery;
        [DataMember] public string InterDest;
        [DataMember] public string ComsumerDest;
    }
}
