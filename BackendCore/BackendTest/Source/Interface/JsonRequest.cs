// Interface 
// * Name    : JsonRequest
// * Version : 0.5
// * Usage   : Request Format from front end to back end
// * Response

using System.Runtime.Serialization;

namespace BackendCore.Source.Interface
{
    //
    // Main Type : Request 
    [DataContract]
    public class JsonRequest
    {
        [DataMember] public int RequestID;
        [DataMember] public JsonReq_CarInfo CarInfo;
        [DataMember] public JsonReq_Shipment Shipment;
        [DataMember] public JsonReq_Cost Cost;
        [DataMember] public JsonReq_Commission Commission;
        [DataMember] public JsonReq_Insurance Insurance;
        [DataMember] public JsonReq_Repair Repair;
    }

    //
    // define elements of JsonRequest.
    [DataContract]
    public class JsonReq_CarInfo
    {
        [DataMember] public string Company;
        [DataMember] public string Model;
        [DataMember] public string Trim;
    }

    [DataContract]
    public class JsonReq_Shipment
    {
        [DataMember] public string SalesType;
        [DataMember] public string Discount;
        [DataMember] public string Delivery;
        [DataMember] public string InterDest;
        [DataMember] public string ComsumerDest;
    }

    [DataContract]
    public class JsonReq_Cost
    {
        [DataMember] public int BasePrice;
        [DataMember] public int OptionPrice;
        [DataMember] public string OptionInfo;
        [DataMember] public int Deposit;
        [DataMember] public int PrePayment;
    }

    [DataContract]
    public class JsonReq_Commission
    {
        [DataMember] public double CMCommission;
        [DataMember] public double AGCommission;
    }

    [DataContract]
    public class JsonReq_Insurance
    {
        [DataMember] public string CorporateType;
        [DataMember] public int DriverAge;
        [DataMember] public int Coverage1;
        [DataMember] public int Coverage2;
        [DataMember] public int Exemption;
    }

    [DataContract]
    public class JsonReq_Repair
    {
        [DataMember] public int RepairType;
    }
}
