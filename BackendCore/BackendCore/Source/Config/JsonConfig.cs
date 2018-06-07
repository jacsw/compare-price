// Interface 
// * Name    : Json Configuration
// * Version : 0.5
// * Usage   : Request Format from front end to back end
// * Response

using System.Runtime.Serialization;

namespace BackendCore.Source.Config
{
    //
    // Main Type : Json Capital File Config
    [DataContract]
    class JsonCapitalList
    {
        [DataMember] public string Capital;
        [DataMember] public string File;
    }

    //
    // Main Type : Json Configuration
    [DataContract]
    class JsonCapitalConfig
    {
        [DataMember] public string CapitalName;
        [DataMember] public string ExcelFile;
        [DataMember] public string Worksheet;

        [DataMember] public JsonConfig_CarInfo CarInfo;
        [DataMember] public JsonConfig_Price Price;
        [DataMember] public JsonConfig_Payment Payment;

        [DataMember] public JsonConfig_ExcelPos[] Extra;
    }

    //
    // define elements of Json Config.
    [DataContract]
    public class JsonConfig_ExcelPos
    {
        [DataMember] public int Row;
        [DataMember] public int Col;
    }

    [DataContract]
    public class JsonConfig_CarInfo
    {
        [DataMember] public JsonConfig_ExcelPos Company;
        [DataMember] public JsonConfig_ExcelPos Model;
        [DataMember] public JsonConfig_ExcelPos Trim;
    }

    [DataContract]
    public class JsonConfig_Price
    {
        [DataMember] public JsonConfig_ExcelPos TotalPrice;
        [DataMember] public JsonConfig_ExcelPos BasePrice;
        [DataMember] public JsonConfig_ExcelPos OptionPrice;
        [DataMember] public JsonConfig_ExcelPos OptionInfo;
    }

    [DataContract]
    public class JsonConfig_Payment
    {
        [DataMember] public JsonConfig_ExcelPos Deposit;
        [DataMember] public JsonConfig_ExcelPos PrePayment;
        [DataMember] public JsonConfig_ExcelPos Duration;
        [DataMember] public JsonConfig_ExcelPos MonthlyFee;
        [DataMember] public JsonConfig_ExcelPos ResidualValue;
        [DataMember] public JsonConfig_ExcelPos ResidualRate;
    }
}
